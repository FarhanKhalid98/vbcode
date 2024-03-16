VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmExpiryDamageClaimInvoice 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmExpiryDamageClaimInvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1950
      Left            =   1733
      TabIndex        =   52
      Top             =   6315
      Width           =   11895
      Begin VB.ComboBox CmbRPackName 
         Height          =   315
         Left            =   4935
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   0
         Width           =   1935
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ReplyGrid 
         Height          =   1635
         Left            =   0
         TabIndex        =   18
         Top             =   315
         Width           =   11895
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   10
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
         stylesets(0).Picture=   "FrmExpiryDamageClaimInvoice.frx":0ECA
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
         Columns.Count   =   10
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "Product ID"
         Columns(0).Name =   "ProductID"
         Columns(0).CaptionAlignment=   2
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3122
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Code"
         Columns(1).CaptionAlignment=   2
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   5583
         Columns(2).Caption=   "Product Name"
         Columns(2).Name =   "ProductName"
         Columns(2).CaptionAlignment=   2
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3413
         Columns(3).Caption=   "Pack Name"
         Columns(3).Name =   "PackName"
         Columns(3).CaptionAlignment=   2
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   1032
         Columns(4).Caption=   "Pack"
         Columns(4).Name =   "Pack"
         Columns(4).Alignment=   1
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   1640
         Columns(5).Caption=   "Qty(P)"
         Columns(5).Name =   "RQtyPack"
         Columns(5).Alignment=   1
         Columns(5).CaptionAlignment=   2
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   4
         Columns(5).FieldLen=   256
         Columns(6).Width=   1588
         Columns(6).Caption=   "Qty(L)"
         Columns(6).Name =   "RQtyLoose"
         Columns(6).Alignment=   1
         Columns(6).CaptionAlignment=   2
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   4
         Columns(6).FieldLen=   256
         Columns(7).Width=   1931
         Columns(7).Caption=   "Price"
         Columns(7).Name =   "Price"
         Columns(7).Alignment=   1
         Columns(7).CaptionAlignment=   2
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   2223
         Columns(8).Caption=   "Amount"
         Columns(8).Name =   "Amount"
         Columns(8).Alignment=   1
         Columns(8).CaptionAlignment=   2
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   5
         Columns(8).FieldLen=   256
         Columns(9).Width=   3200
         Columns(9).Visible=   0   'False
         Columns(9).Caption=   "PackingID"
         Columns(9).Name =   "PackingID"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   20981
         _ExtentY        =   2884
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
      Begin SITextBox.Txt TxtRMultiplier 
         Height          =   315
         Left            =   6870
         TabIndex        =   14
         Top             =   0
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         Alignment       =   1
         Appearance      =   0
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
      Begin SITextBox.Txt TxtRQtyLoose 
         Height          =   315
         Left            =   8385
         TabIndex        =   16
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
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
      Begin SITextBox.Txt TxtRQtyPack 
         Height          =   315
         Left            =   7455
         TabIndex        =   15
         Top             =   0
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         Alignment       =   1
         Appearance      =   0
         MaxLength       =   6
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
      Begin SITextBox.Txt TxtRCode 
         Height          =   315
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1410
         _ExtentX        =   2487
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
      Begin JeweledBut.JeweledButton BtnRProduct 
         Height          =   330
         Left            =   1410
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   0
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
         MICON           =   "FrmExpiryDamageClaimInvoice.frx":0EE6
         BC              =   12632256
         FC              =   0
      End
      Begin SITextBox.Txt TxtRProductName 
         Height          =   315
         Left            =   1770
         TabIndex        =   62
         Top             =   0
         Width           =   3165
         _ExtentX        =   5583
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
      Begin SITextBox.Txt TxtRPrice 
         Height          =   315
         Left            =   9285
         TabIndex        =   17
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
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
      End
      Begin SITextBox.Txt TxtRAmount 
         Height          =   315
         Left            =   10380
         TabIndex        =   63
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
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
   End
   Begin VB.Frame FraClaim 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1950
      Left            =   1733
      TabIndex        =   47
      Top             =   3120
      Width           =   11895
      Begin VB.ComboBox CmbCPackName 
         Height          =   315
         Left            =   3990
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   0
         Width           =   1665
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ClaimGrid 
         Height          =   1635
         Left            =   0
         TabIndex        =   9
         Top             =   315
         Width           =   11895
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   12
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
         stylesets(0).Picture=   "FrmExpiryDamageClaimInvoice.frx":0F02
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
         Columns.Count   =   12
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "Product ID"
         Columns(0).Name =   "ProductID"
         Columns(0).CaptionAlignment=   2
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3122
         Columns(1).Caption=   "Code"
         Columns(1).Name =   "Code"
         Columns(1).CaptionAlignment=   2
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3889
         Columns(2).Caption=   "Product Name"
         Columns(2).Name =   "ProductName"
         Columns(2).CaptionAlignment=   2
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   2963
         Columns(3).Caption=   "Pack Name"
         Columns(3).Name =   "PackName"
         Columns(3).CaptionAlignment=   2
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   953
         Columns(4).Caption=   "Pack"
         Columns(4).Name =   "Pack"
         Columns(4).Alignment=   1
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   1323
         Columns(5).Caption=   "EQty(P)"
         Columns(5).Name =   "EQtyPack"
         Columns(5).Alignment=   1
         Columns(5).CaptionAlignment=   2
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   4
         Columns(5).FieldLen=   256
         Columns(6).Width=   1429
         Columns(6).Caption=   "EQty(L)"
         Columns(6).Name =   "EQtyLoose"
         Columns(6).Alignment=   1
         Columns(6).CaptionAlignment=   2
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   4
         Columns(6).FieldLen=   256
         Columns(7).Width=   1323
         Columns(7).Caption=   "DQty(P)"
         Columns(7).Name =   "DQtyPack"
         Columns(7).Alignment=   1
         Columns(7).CaptionAlignment=   2
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   4
         Columns(7).FieldLen=   256
         Columns(8).Width=   1429
         Columns(8).Caption=   "DQty(L)"
         Columns(8).Name =   "DQtyLoose"
         Columns(8).Alignment=   1
         Columns(8).CaptionAlignment=   2
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   4
         Columns(8).FieldLen=   256
         Columns(9).Width=   1138
         Columns(9).Caption=   "Cost"
         Columns(9).Name =   "Cost"
         Columns(9).Alignment=   1
         Columns(9).CaptionAlignment=   2
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   4
         Columns(9).FieldLen=   256
         Columns(10).Width=   2937
         Columns(10).Caption=   "Amount"
         Columns(10).Name=   "Amount"
         Columns(10).Alignment=   1
         Columns(10).CaptionAlignment=   2
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   5
         Columns(10).FieldLen=   256
         Columns(11).Width=   3200
         Columns(11).Visible=   0   'False
         Columns(11).Caption=   "PackingID"
         Columns(11).Name=   "PackingID"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   20981
         _ExtentY        =   2884
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
      Begin SITextBox.Txt TxtCMultiplier 
         Height          =   315
         Left            =   5655
         TabIndex        =   4
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         Alignment       =   1
         Appearance      =   0
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
      Begin SITextBox.Txt TxtCEQtyLoose 
         Height          =   315
         Left            =   6945
         TabIndex        =   6
         Top             =   0
         Width           =   810
         _ExtentX        =   1429
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
      Begin SITextBox.Txt TxtCEQtyPack 
         Height          =   315
         Left            =   6195
         TabIndex        =   5
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         Alignment       =   1
         Appearance      =   0
         MaxLength       =   6
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
      Begin SITextBox.Txt TxtCCode 
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1410
         _ExtentX        =   2487
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
      Begin JeweledBut.JeweledButton BtnCProduct 
         Height          =   330
         Left            =   1410
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   0
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
         MICON           =   "FrmExpiryDamageClaimInvoice.frx":0F1E
         BC              =   12632256
         FC              =   0
      End
      Begin SITextBox.Txt TxtCProductName 
         Height          =   315
         Left            =   1770
         TabIndex        =   49
         Top             =   0
         Width           =   2220
         _ExtentX        =   3916
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
      Begin SITextBox.Txt TxtCDQtyLoose 
         Height          =   315
         Left            =   8505
         TabIndex        =   8
         Top             =   0
         Width           =   810
         _ExtentX        =   1429
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
      Begin SITextBox.Txt TxtCDQtyPack 
         Height          =   315
         Left            =   7755
         TabIndex        =   7
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         Alignment       =   1
         Appearance      =   0
         MaxLength       =   6
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
      Begin SITextBox.Txt TxtCAmount 
         Height          =   315
         Left            =   9960
         TabIndex        =   50
         Top             =   0
         Width           =   1920
         _ExtentX        =   3387
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
      Begin SITextBox.Txt TxtCCost 
         Height          =   315
         Left            =   9315
         TabIndex        =   51
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
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
   End
   Begin SITextBox.Txt TxtClaimID 
      Height          =   315
      Left            =   1898
      TabIndex        =   0
      Top             =   2055
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
      Left            =   9053
      TabIndex        =   23
      Top             =   9210
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
      MICON           =   "FrmExpiryDamageClaimInvoice.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7717
      TabIndex        =   20
      Top             =   9210
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
      MICON           =   "FrmExpiryDamageClaimInvoice.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5045
      TabIndex        =   22
      Top             =   9210
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
      MICON           =   "FrmExpiryDamageClaimInvoice.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10392
      TabIndex        =   24
      Top             =   9210
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
      MICON           =   "FrmExpiryDamageClaimInvoice.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6381
      TabIndex        =   21
      Top             =   9210
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
      MICON           =   "FrmExpiryDamageClaimInvoice.frx":0FAA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtClaimTotalAmount 
      Height          =   315
      Left            =   12233
      TabIndex        =   27
      Top             =   5340
      Width           =   1215
      _ExtentX        =   2143
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpClaimDate 
      Height          =   315
      Left            =   3188
      TabIndex        =   1
      Top             =   2055
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   11633
      TabIndex        =   31
      Top             =   1740
      Visible         =   0   'False
      Width           =   1860
      _ExtentX        =   3281
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpReplyDate 
      Height          =   315
      Left            =   1778
      TabIndex        =   10
      Top             =   5610
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   5513
      TabIndex        =   11
      Tag             =   "NC"
      Top             =   5610
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
      Left            =   6548
      TabIndex        =   64
      Tag             =   "NC"
      Top             =   5610
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
      Left            =   6188
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   5610
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
      MICON           =   "FrmExpiryDamageClaimInvoice.frx":0FC6
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnReturnAll 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   3623
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   5595
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   582
      TX              =   "Return All"
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
      MICON           =   "FrmExpiryDamageClaimInvoice.frx":0FE2
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtReplyTotalAmount 
      Height          =   315
      Left            =   12263
      TabIndex        =   69
      Top             =   8610
      Width           =   1215
      _ExtentX        =   2143
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
   Begin SITextBox.Txt TxtRepliedAmount 
      Height          =   315
      Left            =   7178
      TabIndex        =   19
      Top             =   8610
      Width           =   1215
      _ExtentX        =   2143
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
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3709
      TabIndex        =   72
      Top             =   9210
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
      MICON           =   "FrmExpiryDamageClaimInvoice.frx":0FFE
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationName 
      Height          =   315
      Left            =   6105
      TabIndex        =   73
      Tag             =   "NC"
      Top             =   1995
      Width           =   1665
      _ExtentX        =   2937
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
      Left            =   5745
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   1995
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
      MICON           =   "FrmExpiryDamageClaimInvoice.frx":101A
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   5040
      TabIndex        =   75
      Tag             =   "NC"
      Top             =   1995
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
   Begin VB.Label LblOrganizationID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   5040
      TabIndex        =   77
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label LblOrganizationName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   6225
      TabIndex        =   76
      Top             =   1800
      Width           =   1350
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Replied Amount"
      Height          =   195
      Left            =   7178
      TabIndex        =   71
      Top             =   8385
      Width           =   1125
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Left            =   12263
      TabIndex        =   70
      Top             =   8385
      Width           =   945
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   5513
      TabIndex        =   67
      Top             =   5415
      Width           =   585
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   6548
      TabIndex        =   66
      Top             =   5415
      Width           =   840
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3503
      TabIndex        =   60
      Top             =   6090
      Width           =   1020
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   1733
      TabIndex        =   59
      Top             =   6090
      Width           =   375
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   11048
      TabIndex        =   58
      Top             =   6090
      Width           =   360
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   8603
      TabIndex        =   57
      Top             =   6090
      Width           =   375
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   6683
      TabIndex        =   56
      Top             =   6090
      Width           =   840
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
      Height          =   195
      Left            =   10193
      TabIndex        =   55
      Top             =   6090
      Width           =   465
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
      Height          =   195
      Left            =   9248
      TabIndex        =   54
      Top             =   6090
      Width           =   480
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   12338
      TabIndex        =   53
      Top             =   6090
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   11048
      TabIndex        =   46
      Top             =   2895
      Width           =   315
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reply Date"
      Height          =   195
      Left            =   1778
      TabIndex        =   45
      Top             =   5415
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "------- Damage --------"
      Height          =   195
      Left            =   9518
      TabIndex        =   44
      Top             =   2625
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
      Height          =   195
      Left            =   10268
      TabIndex        =   43
      Top             =   2895
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
      Height          =   195
      Left            =   9563
      TabIndex        =   42
      Top             =   2895
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "-------- Expiry --------"
      Height          =   195
      Left            =   7943
      TabIndex        =   41
      Top             =   2625
      Width           =   1230
   End
   Begin VB.Label LblStock 
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   450
      Left            =   10058
      TabIndex        =   40
      Top             =   1815
      Width           =   975
   End
   Begin VB.Label LblStockCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   10058
      TabIndex        =   39
      Top             =   1500
      Width           =   720
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry / Damage Claim Invoice"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   38
      Top             =   270
      Width           =   4890
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   11723
      TabIndex        =   37
      Top             =   2895
      Width           =   540
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
      Height          =   195
      Left            =   8033
      TabIndex        =   36
      Top             =   2895
      Width           =   480
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
      Height          =   195
      Left            =   8753
      TabIndex        =   35
      Top             =   2895
      Width           =   465
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   5723
      TabIndex        =   34
      Top             =   2895
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   7388
      TabIndex        =   33
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   11633
      TabIndex        =   32
      Top             =   1545
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   1733
      TabIndex        =   30
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3503
      TabIndex        =   29
      Top             =   2895
      Width           =   1020
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Left            =   12233
      TabIndex        =   28
      Top             =   5115
      Width           =   945
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Claim Date"
      Height          =   195
      Left            =   3203
      TabIndex        =   26
      Top             =   1860
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Claim ID"
      Height          =   195
      Left            =   1913
      TabIndex        =   25
      Top             =   1860
      Width           =   585
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmExpiryDamageClaimInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vCounter As Integer
Dim RsBodyClaim As New ADODB.Recordset
Dim RsBodyReply As New ADODB.Recordset
Dim Flag As Boolean
Dim ssql As String
Dim vStrSQL As String
Dim vQtyLoose As Double

'----------------------------------

Private Sub SubCalculateBodyClaim()
   TxtCAmount.Text = (Val(TxtCCost.Text) * (Val(TxtCEQtyPack.Text) * Val(TxtCMultiplier.Text) + Val(TxtCEQtyLoose.Text))) + (Val(TxtCCost.Text) * (Val(TxtCDQtyPack.Text) * Val(TxtCMultiplier.Text) + Val(TxtCDQtyLoose.Text)))
End Sub

Private Sub SubCalculateBodyReply()
   TxtRAmount.Text = (Val(TxtRPrice.Text) * (Val(TxtRQtyPack.Text) * Val(TxtRMultiplier.Text) + Val(TxtRQtyLoose.Text)))
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
    With cn.Execute(vStrSQL)
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

Private Function FunSelectClaimProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectClaimProduct = False: Exit Function
      TxtCCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtCCode.Text) = "" Then Exit Function
    If Len(TxtCCode.Text) <= 5 Then
      TxtCCode.Text = Right("00000" + CStr(Val(TxtCCode.Text)), 5)
    End If
    If TxtCCode.Text = "" Then FunSelectClaimProduct = False: Exit Function
    vStrSQL = " SELECT p.productid, Code, ProductName, PackingName, isnull(Multiplier,0) as Multiplier " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " where p.productid = '" & TxtCCode.Text & "' or code='" & TxtCCode.Text & "'"

   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtCProductName.Text = !ProductName
         TxtCMultiplier.Text = IIf(!Multiplier = 0, "", !Multiplier)
         If Not IsNull(!PackingName) Then CmbCPackName.Text = !PackingName
         
         vStrSQL = "select isnull(dbo.FunStock('" & TxtCCode.Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpClaimDate.DateValue + 1 & "',0),0)"
         vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtCCode.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbCPackName.Text
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtCCode.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
         
'         With CN.Execute("select * from CurrentStockExpiry where productid ='" & TxtProductID.Text & "'")
'            If .RecordCount > 0 Then
'               TxtCCost.Text = !Cost
'               'LblStock.Caption = !QtyLoose
'            Else
'               TxtCCost.Text = ""
'               'LblStock.Caption = 0
'            End If
'         End With
         'LblStock.Visible = True
         'LblStockCaption.Visible = True
         SubCalculateBodyClaim
'         Char.Speak TxtCProductName.Text
         FunSelectClaimProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectClaimProduct = False
         .Close
         TxtProductID.Text = ""
         TxtCCode.Text = ""
         If CmbCPackName.ListCount > 0 Then CmbCPackName.ListIndex = 0
         TxtCProductName.Text = ""
         TxtCMultiplier.Text = ""
         TxtCCost.Text = ""
         TxtCAmount.Text = ""
         LblStock.Visible = False
         LblStockCaption.Visible = False
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectReplyProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectReplyProduct = False: Exit Function
      TxtRCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtRCode.Text) = "" Then Exit Function
    If Len(TxtRCode.Text) <= 5 Then
      TxtRCode.Text = Right("00000" + CStr(Val(TxtRCode.Text)), 5)
    End If
    If TxtRCode.Text = "" Then FunSelectReplyProduct = False: Exit Function
    vStrSQL = " SELECT p.productid, Code, ProductName, PackingName, isnull(Multiplier,0) as Multiplier, PurPrice" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " where p.productid = '" & TxtRCode.Text & "' or code='" & TxtRCode.Text & "'"

   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtRProductName.Text = !ProductName
         TxtRMultiplier.Text = IIf(!Multiplier = 0, "", !Multiplier)
         TxtRPrice.Text = !PurPrice
         If Not IsNull(!PackingName) Then CmbRPackName.Text = !PackingName
         vStrSQL = "select isnull(dbo.FunStock('" & TxtRCode.Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpReplyDate.DateValue + 1 & "',0),0)"
         vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtRCode.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbCPackName.Text
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtRCode.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
'         With CN.Execute("select cost from currentstock where productid ='" & TxtProductID.Text & "'")
'            If .RecordCount > 0 Then
'               TxtCCost.Text = !Cost
'            Else
'               TxtCCost.Text = ""
'            End If
'         End With
'         With CN.Execute("select qtyloose from currentstockStore where productid ='" & TxtProductID.Text & "' and storeid = " & TxtStoreID.Text)
'            If .RecordCount > 0 Then
'               LblStock.Caption = !QtyLoose
'            Else
'               LblStock.Caption = 0
'            End If
'         End With
         SubCalculateBodyReply
'         Char.Speak TxtRProductName.Text
         FunSelectReplyProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectReplyProduct = False
         .Close
         TxtProductID.Text = ""
         TxtRCode.Text = ""
         If CmbRPackName.ListCount > 0 Then CmbRPackName.ListIndex = 0
         TxtRProductName.Text = ""
         TxtRMultiplier.Text = ""
         TxtRPrice.Text = ""
         TxtRAmount.Text = ""
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
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
   
    ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniExpiryDamageClaimInvoice", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
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
   ReplyGrid.Redraw = False
   ReplyGrid.MoveFirst
   For vCounter = 1 To ReplyGrid.Rows
      If Trim(ReplyGrid.Columns("ProductID").Text) <> "" Then
         cn.Execute "Delete from ExpiryReplyBody where ReplyID = " & Val(TxtClaimID.Text) & " and ProductID='" & ReplyGrid.Columns("Productid").Text & "'"
      End If
      ReplyGrid.MoveNext
   Next vCounter
   ReplyGrid.RemoveAll
   ReplyGrid.Redraw = True
   cn.Execute "Delete from ExpiryReplyHeader where ReplyID = " & Val(TxtClaimID.Text)
   ClaimGrid.Redraw = False
   ClaimGrid.MoveFirst
   For vCounter = 1 To ClaimGrid.Rows
      If Trim(ClaimGrid.Columns("ProductID").Text) <> "" Then
         cn.Execute "Delete from ExpiryClaimsBody where ClaimID = " & Val(TxtClaimID.Text) & " and ProductID='" & ClaimGrid.Columns("Productid").Text & "'"
      End If
      ClaimGrid.MoveNext
   Next vCounter
   ClaimGrid.RemoveAll
   cn.Execute "Delete from ExpiryClaimsHeader where ClaimID = " & Val(TxtClaimID.Text)
   ClaimGrid.Redraw = True
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   ReplyGrid.Redraw = True
   ClaimGrid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchClaim.Show vbModal
   If SchClaim.ParaOutClaimID <> 0 Then
      TxtClaimID.Text = SchClaim.ParaOutClaimID
      GetClaim
   End If
End Sub

'Private Sub BtnPrint_Click()
'On Error GoTo ErrorHandler
'   vStrSql = "select u.username, h.ClaimID, h.ClaimDate, h.TotalAmount as tbill, isnull(h.discount,0) as discount, isnull(h.cashReceived,0) as cashReceived, p.productname, b.qty, b.price-b.discountvalue as price, b.amount" _
'            + " from ExpiryClaimsHeader h inner join ExpiryClaimsBody b on h.ClaimID = b.ClaimID and h.ClaimDate = b.ClaimDate" _
'            + " inner join products p on p.productid = b.productid" _
'            + " inner join users u on u.UserNo = h.UserNo" _
'            + " where h.ClaimID= " & Val(TxtClaimID.Text) & " and h.ClaimDate='" & DtpClaimDate.DateValue & "' order by SerialNo"
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

Private Sub BtnCProduct_Click()
   If FunSelectClaimProduct(ssButton, True) = True Then
      CmbCPackName.SetFocus
   Else
      TxtCCode.SetFocus
   End If
End Sub

Private Sub BtnPrint_Click()
  If FunRefreshData = False Then Exit Sub
  If RsBodyClaim.RecordCount = 0 And RsBodyReply.RecordCount = 0 Then
    MsgBox "No record found", vbInformation, "Information"
    Exit Sub
  ElseIf RsBodyReply.RecordCount = 0 Then
    Call SetCrystalReport1
    Call SetReportParameterField
    RptReportViewer.Show vbModal, Me
  Else
    Call SetCrystalReport
    Call SetReportParameterField
    RptReportViewer.Show vbModal, Me
  End If
  Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnRProduct_Click()
   If FunSelectReplyProduct(ssButton, True) = True Then
      CmbRPackName.SetFocus
   Else
      TxtRCode.SetFocus
   End If
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniExpiryDamageClaimInvoice", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
  If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
    MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
    Exit Sub
  End If
   '''''''''''''''''''''''Check Posing Date'''''''''''''''''''''''''''''''''
    vStrSQL = "Select isnull(max(EntryDate),'01/01/1990') from AdminClosing where touserno = " & vUser & " and Entrydate <='" & Date & "'"
    With cn.Execute(vStrSQL)
        If .Fields(0).Value >= DtpClaimDate.DateValue Then
            MsgBox "Data can not be saved in back date of posting Date ( " & Format(.Fields(0).Value, "dd/mm/yyyy") & " )", vbInformation, Me.Caption
            Exit Sub
        End If
    End With
   '''''''''''''''''''''''Check Entry Date'''''''''''''''''''''''''''''''''
    If ObjRegistry.isEntryDate = True Then
       If ObjRegistry.FromDate > Date Or ObjRegistry.ToDate < Date Then
         MsgBox "Data can not be saved Because Date is not set according to the Software's Entry date", vbInformation, Me.Caption
         Exit Sub
       End If
    End If
'  Header Validation
   If DtpClaimDate.Enabled Then
      If cn.Execute("Select * from ExpiryClaimsHeader where ClaimID = " & Val(TxtClaimID.Text) & " and ClaimDate = '" & DtpClaimDate.DateValue & "'").RecordCount > 0 Then
         'MsgBox "This Bill ID already exists. A new Bill ID. has been generated. Please try again", vbCritical, "Alert"
         TxtClaimID.Text = FunGetMaxID
         'Exit Sub
      End If
   End If
   ''''''''''''''''''''''''Check Organization '''''''''''''''''''''''''''''''''
  If ObjRegistry.OrganizationMandatory = True And TxtOrganizationID.Text = "" Then
    MsgBox "Please Select Organization", vbInformation, Me.Caption
    If TxtOrganizationID.Visible = True Then TxtOrganizationID.SetFocus
    Exit Sub
  End If
  
  RsBodyClaim.Filter = 0
  If RsBodyClaim.RecordCount = 0 Then
      MsgBox "Please enter at least one product to Claim", vbExclamation, "Alert"
      If TxtCCode.Visible And TxtCCode.Enabled Then TxtCCode.SetFocus
      Exit Sub
  End If
  'Body Validation
  ' validation has been performed when a row is added to the ClaimGrid

  'Saving record
   cn.BeginTrans
   ssql = "select * from ExpiryClaimsHeader where ClaimID=" & Val(TxtClaimID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, cn, adOpenStatic, adLockPessimistic
      If .BOF Then
         .AddNew
         !ClaimID = Val(TxtClaimID.Text)
      End If
'      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !ClaimDate = DtpClaimDate.DateValue
      !TotalAmount = Round(Val(TxtClaimTotalAmount.Text))
      !UserNo = vUser
      !SessionID = IIf(Trim(vSessionID) = 0, Null, Val(vSessionID))
      .Update
      .Close
   End With
   RsBodyReply.Filter = 0
   ssql = "select * from ExpiryReplyHeader where ReplyID=" & Val(TxtClaimID.Text)
   With Rs
      .Open ssql, cn, adOpenStatic, adLockPessimistic
      If RsBodyReply.RecordCount = 0 Then
         If TxtClaimID.Enabled = False Then
            'delete reply body
            If .RecordCount = 1 Then
               'CN.Execute "Delete from ExpiryReplyBody where ReplyID = " & Val(TxtClaimID.Text)
               RsBodyReply.UpdateBatch
               cn.Execute "Delete from ExpiryReplyHeader where ReplyID = " & Val(TxtClaimID.Text)
            End If
         End If
      Else
         If .BOF Then
            .AddNew
            !ReplyID = Val(TxtClaimID.Text)
         End If
         !ReplyDate = DtpReplyDate.DateValue
         !TotalAmount = Round(Val(TxtReplyTotalAmount.Text))
         !RepliedAmount = IIf(Val(TxtRepliedAmount.Text) = 0, Null, TxtRepliedAmount.Text)
         !StoreID = TxtStoreID.Text
         !UserNo = vUser
         .Update
         .Close
         RsBodyReply.Filter = 0
         RsBodyReply.MoveFirst
         For vCounter = 1 To RsBodyReply.RecordCount
            RsBodyReply!ReplyID = Val(TxtClaimID.Text)
            RsBodyReply.MoveNext
         Next vCounter
         RsBodyReply.UpdateBatch
      End If
   End With
   With RsBodyClaim
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !ClaimID = Val(TxtClaimID.Text)
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   ClaimGrid.Redraw = True
   ReplyGrid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToClaim()
   RsBodyClaim.Filter = 0
   If RsBodyClaim.State = adStateOpen Then RsBodyClaim.Close
   RsBodyClaim.Open "Select * from ExpiryClaimsBody where ClaimID=" & Val(TxtClaimID.Text), cn, adOpenStatic, adLockBatchOptimistic
   If RsBodyClaim.RecordCount > 0 Then
      ssql = "select p.productname, code, b.* from ExpiryClaimsBody b join products p on p.productid = b.productid where ClaimID=" & Val(TxtClaimID.Text)
      With cn.Execute(ssql)
         ClaimGrid.Redraw = False
         ClaimGrid.MoveFirst
         ClaimGrid.RemoveAll
         ClaimGrid.AllowAddNew = True
         TxtClaimTotalAmount.Text = 0
         While Not .EOF
            ClaimGrid.AddNew
            ClaimGrid.Columns("ProductID").Text = !Productid
            ClaimGrid.Columns("Code").Text = !Code
            ClaimGrid.Columns("ProductName").Text = !ProductName
            If !PackingID = 0 Or IsNull(!PackingID) Then
               ClaimGrid.Columns("PackingID").Value = ""
            Else
               ClaimGrid.Columns("PackingID").Value = !PackingID
            End If
            If !PackingID = 0 Or IsNull(!PackingID) Then
               ClaimGrid.Columns("PackName").Text = ""
            Else
               ClaimGrid.Columns("PackName").Text = cn.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
            End If
            ClaimGrid.Columns("Pack").Value = !Multiplier
            ClaimGrid.Columns("EQtyPack").Value = !EQtyPack
            ClaimGrid.Columns("EQtyLoose").Value = !EQtyLoose
            ClaimGrid.Columns("DQtyPack").Value = !DQtyPack
            ClaimGrid.Columns("DQtyLoose").Value = !DQtyLoose
            ClaimGrid.Columns("Cost").Value = !Cost
            ClaimGrid.Columns("Amount").Value = !Amount
            TxtClaimTotalAmount.Text = Val(TxtClaimTotalAmount.Text) + Val(!Amount)
            .MoveNext
         Wend
         .Close
      End With
      ClaimGrid.AddNew
      ClaimGrid.Columns("Code").Text = " "
      ClaimGrid.AllowAddNew = False
      ClaimGrid.Redraw = True
   End If
End Sub

Private Sub PopulateDataToReply()
   RsBodyReply.Filter = 0
   If RsBodyReply.State = adStateOpen Then RsBodyReply.Close
   RsBodyReply.Open "Select * from ExpiryReplyBody where ReplyID=" & Val(TxtClaimID.Text), cn, adOpenStatic, adLockBatchOptimistic
   If RsBodyReply.RecordCount > 0 Then
      ssql = "select p.productname, code, b.* from ExpiryReplyBody b join products p on p.productid = b.productid where ReplyID=" & Val(TxtClaimID.Text)
      With cn.Execute(ssql)
         ReplyGrid.Redraw = False
         ReplyGrid.MoveFirst
         ReplyGrid.RemoveAll
         ReplyGrid.AllowAddNew = True
         TxtReplyTotalAmount.Text = 0
         While Not .EOF
            ReplyGrid.AddNew
            ReplyGrid.Columns("ProductID").Text = !Productid
            ReplyGrid.Columns("Code").Text = !Code
            ReplyGrid.Columns("ProductName").Text = !ProductName
            If !PackingID = 0 Or IsNull(!PackingID) Then
               ReplyGrid.Columns("PackingID").Value = ""
            Else
               ReplyGrid.Columns("PackingID").Value = !PackingID
            End If
            If !PackingID = 0 Or IsNull(!PackingID) Then
               ReplyGrid.Columns("PackName").Text = ""
            Else
               ReplyGrid.Columns("PackName").Text = cn.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
            End If
            ReplyGrid.Columns("Pack").Value = !Multiplier
            ReplyGrid.Columns("RQtyPack").Value = !QtyPack
            ReplyGrid.Columns("RQtyLoose").Value = !QtyLoose
            ReplyGrid.Columns("Price").Value = !Cost
            ReplyGrid.Columns("Amount").Value = !Amount
            TxtReplyTotalAmount.Text = Val(TxtReplyTotalAmount.Text) + Val(!Amount)
            .MoveNext
         Wend
         .Close
      End With
      ReplyGrid.AddNew
      ReplyGrid.Columns("Code").Text = " "
      ReplyGrid.AllowAddNew = False
      ReplyGrid.Redraw = True
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
      Call PopulateDataToClaim
      Call PopulateDataToReply
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      LblStock.Visible = False
      LblStockCaption.Visible = False
      TxtClaimID.Text = FunGetMaxID()
      DtpClaimDate.Enabled = True
      If DtpClaimDate.Enabled And DtpClaimDate.Visible Then DtpClaimDate.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      DtpClaimDate.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      LblStock.Visible = False
      LblStockCaption.Visible = False
      TxtCCode.Enabled = True
      BtnCProduct.Enabled = True
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

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      If TxtRCode.Enabled Then TxtRCode.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub CmbCPackName_Click()
   If CmbCPackName.Text = "" Then
      TxtCMultiplier.Enabled = False
      TxtCEQtyPack.Enabled = False
      TxtCDQtyPack.Enabled = False
      TxtCMultiplier.Text = ""
      TxtCEQtyPack.Text = ""
      TxtCDQtyPack.Text = ""
   Else
      TxtCMultiplier.Enabled = True
      TxtCEQtyPack.Enabled = True
      TxtCDQtyPack.Enabled = True
   End If
End Sub

Private Sub CmbRPackName_Click()
   If CmbRPackName.Text = "" Then
      TxtRMultiplier.Enabled = False
      TxtRQtyPack.Enabled = False
      TxtRMultiplier.Text = ""
      TxtRQtyPack.Text = ""
   Else
      TxtRMultiplier.Enabled = True
      TxtRQtyPack.Enabled = True
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name = ClaimGrid.Name Then
         ClaimGrid_DblClick
      ElseIf ActiveControl.Name = ReplyGrid.Name Then
         ReplyGrid_DblClick
      Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   'ElseIf KeyCode = vbKeyEscape Then
   '   Call SubClearDetailArea: TxtCCode.SetFocus
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
         Case vbKeyR
            If BtnDelete.Enabled Then BtnDelete_Click
            KeyCode = 0
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCCode.Name: If FunSelectClaimProduct(ssFunctionKey, True) = True Then CmbCPackName.SetFocus
         Case TxtRCode.Name: If FunSelectReplyProduct(ssFunctionKey, True) = True Then CmbRPackName.SetFocus
      End Select
   ElseIf ActiveControl.Name = TxtCCode.Name Then
      If KeyCode = vbKeyDown Then
         ClaimGrid.SetFocus
      ElseIf KeyCode = vbKeyF12 Then
         KeyCode = 0
         DtpReplyDate.SetFocus
      End If
   ElseIf ActiveControl.Name = TxtRCode.Name Then
      If KeyCode = vbKeyDown Then
         ReplyGrid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtRCode.Name Then
         KeyCode = 0
         TxtRepliedAmount.SetFocus
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

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Search"

   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   
   DtpClaimDate.DateValue = IIf(Format(Now, "hh") >= Val(ObjRegistry.HourDifference), Date, DateAdd("d", -1, Date))

   TxtStoreID.Text = IIf((ObjRegistry.StoreID = ""), "", ObjRegistry.StoreID)
   FunSelectStore ssValidate, True
   LblStoreID.Visible = ObjRegistry.StoreVisible
   LblStoreName.Visible = ObjRegistry.StoreVisible
   TxtStoreID.Visible = ObjRegistry.StoreVisible
   TxtStoreName.Visible = ObjRegistry.StoreVisible
   BtnStore.Visible = ObjRegistry.StoreVisible
   
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible

'   TxtOrganizationID.Text = ObjRegistry.OrganizationID
'   FunSelectOrganization ssValidate, True
'   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
'   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
'   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
'   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
'   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible
   
   With cn.Execute("Select * from Packings")
      CmbCPackName.AddItem ""
      CmbRPackName.AddItem ""
      While Not .EOF
         CmbCPackName.AddItem !PackingName
         CmbCPackName.ItemData(CmbCPackName.NewIndex) = !PackingID
         CmbRPackName.AddItem !PackingName
         CmbRPackName.ItemData(CmbRPackName.NewIndex) = !PackingID
         .MoveNext
      Wend
      .Close
   End With
   CmbCPackName.ListIndex = 0
   CmbRPackName.ListIndex = 0
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   'If DtpClaimDate.IsDateValid = False Then Exit Function
   FunGetMaxID = cn.Execute("Select isnull(max(ClaimID),0)+1 from ExpiryClaimsHeader").Fields(0)
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
   TxtClaimTotalAmount.Text = 0
   ClaimGrid.CancelUpdate
   ClaimGrid.RemoveAll
   ClaimGrid.AddNew
   ClaimGrid.Columns("Code").Text = " "
   ClaimGrid.Update
   TxtReplyTotalAmount.Text = 0
   ReplyGrid.CancelUpdate
   ReplyGrid.RemoveAll
   ReplyGrid.AddNew
   ReplyGrid.Columns("Code").Text = " "
   ReplyGrid.Update
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
    Set RsBodyClaim = Nothing
    Set RsBodyReply = Nothing
    Set FrmExpiryDamageClaimInvoice = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ClaimGrid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtClaimTotalAmount.Text = Val(TxtClaimTotalAmount.Text) - ClaimGrid.Columns("Amount").Value
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ClaimGrid_DblClick()
   Call ClaimGrid_LostFocus
End Sub

Private Sub ClaimGrid_GotFocus()
   Flag = True
   TxtCCode.Enabled = False
   BtnCProduct.Enabled = False
   'TxtCCode.BackColor = TxtCProductName.BackColor
   'TxtCCode.TabStop = False
End Sub

Private Sub ClaimGrid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub ClaimGrid_LostFocus()
   Flag = False
   If Trim(ClaimGrid.Columns("Code").Text) = "" Then
      TxtCCode.Text = ""
      TxtCCode.Enabled = True
      BtnCProduct.Enabled = True
      TxtCCode.SetFocus
   Else
      TxtCCode.Enabled = False
      BtnCProduct.Enabled = False
      CmbCPackName.SetFocus
      If BtnSave.Enabled = False Then BtnSave.Enabled = True
   End If
End Sub

Private Sub ClaimGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(ClaimGrid.Columns("Code").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub ClaimGrid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromClaimGridToTexBoxes
End Sub

Private Sub ReplyGrid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtReplyTotalAmount.Text = Val(TxtReplyTotalAmount.Text) - ReplyGrid.Columns("Amount").Value
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ReplyGrid_DblClick()
   Call ReplyGrid_LostFocus
End Sub

Private Sub ReplyGrid_GotFocus()
   Flag = True
   TxtRCode.Enabled = False
   BtnRProduct.Enabled = False
   'TxtRCode.BackColor = TxtRProductName.BackColor
   'TxtRCode.TabStop = False
End Sub

Private Sub ReplyGrid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub ReplyGrid_LostFocus()
   Flag = False
   If Trim(ReplyGrid.Columns("Code").Text) = "" Then
      TxtRCode.Text = ""
      TxtRCode.Enabled = True
      BtnRProduct.Enabled = True
      TxtRCode.SetFocus
   Else
      TxtRCode.Enabled = False
      BtnRProduct.Enabled = False
      CmbRPackName.SetFocus
      If BtnSave.Enabled = False Then BtnSave.Enabled = True
   End If
End Sub

Private Sub ReplyGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(ReplyGrid.Columns("Code").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub ReplyGrid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromReplyGridToTexBoxes
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If ActiveControl.Name = ClaimGrid.Name Then
      If Trim(ClaimGrid.Columns("Code").Text) = "" Then Exit Sub
      RsBodyClaim.Filter = "Code='" & TxtCCode.Text & "'"
      If RsBodyClaim.RecordCount > 0 Then RsBodyClaim.Delete
      ClaimGrid.SelBookmarks.RemoveAll
      ClaimGrid.SelBookmarks.Add ClaimGrid.Bookmark
      ClaimGrid.DeleteSelected
      ClaimGrid.Refresh
      RsBodyClaim.Filter = 0
      ClaimGrid.MoveLast
      GetDataBackFromClaimGridToTexBoxes
   End If
   If ActiveControl.Name = ReplyGrid.Name Then
      If Trim(ReplyGrid.Columns("Code").Text) = "" Then Exit Sub
      RsBodyReply.Filter = "Code='" & TxtRCode.Text & "'"
      If RsBodyReply.RecordCount > 0 Then RsBodyReply.Delete
      ReplyGrid.SelBookmarks.RemoveAll
      ReplyGrid.SelBookmarks.Add ReplyGrid.Bookmark
      ReplyGrid.DeleteSelected
      ReplyGrid.Refresh
      RsBodyReply.Filter = 0
      ReplyGrid.MoveLast
      GetDataBackFromReplyGridToTexBoxes
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToClaimGrid()
   Dim vrowcounter As Integer
   If Trim(TxtCCode.Text) = "" Then
      TxtCCode.SetFocus
      Exit Sub
   End If
   If CmbCPackName.ListIndex > 0 Then
      If Trim(TxtCMultiplier.Text) = 0 Then
         TxtCMultiplier.SetFocus
         Exit Sub
      End If
   End If
   If Trim(TxtCEQtyPack.Text) = "" And Trim(TxtCEQtyLoose.Text) = "" And Trim(TxtCDQtyPack.Text) = "" And Trim(TxtCDQtyLoose.Text) = "" Then
      TxtCEQtyLoose.SetFocus
      Exit Sub
   End If
On Error GoTo ErrorHandler
   RsBodyClaim.Filter = "Code='" & TxtCCode.Text & "'"
   If TxtCCode.Enabled Then
      If RsBodyClaim.RecordCount = 0 Then
         RsBodyClaim.AddNew
         ClaimGrid.Columns("ProductID").Text = TxtProductID.Text
         ClaimGrid.Columns("Code").Text = TxtCCode.Text
         RsBodyClaim!Productid = TxtProductID.Text
         RsBodyClaim!Code = TxtCCode.Text
      Else
         ClaimGrid.Redraw = False
         ClaimGrid.MoveFirst
            For vrowcounter = 1 To ClaimGrid.Rows
               If ClaimGrid.Columns("Code").Text = TxtProductID.Text Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailAreaclaim
                  TxtCEQtyLoose.Text = Val(TxtCEQtyLoose.Text) + ClaimGrid.Columns("EQtyLoose").Value
                  TxtCEQtyPack.Text = Val(TxtCEQtyPack.Text) + ClaimGrid.Columns("EQtyPack").Value
                  TxtCDQtyLoose.Text = Val(TxtCDQtyLoose.Text) + ClaimGrid.Columns("DQtyLoose").Value
                  TxtCDQtyPack.Text = Val(TxtCDQtyPack.Text) + ClaimGrid.Columns("DQtyPack").Value
                  TxtClaimTotalAmount.Text = Val(TxtClaimTotalAmount.Text) + Val(TxtCAmount.Text) - Val(ClaimGrid.Columns("Amount").Text)
                  ClaimGrid.Columns("ProductName").Text = TxtCProductName.Text
                  ClaimGrid.Columns("PackName").Text = CmbCPackName.Text
                  ClaimGrid.Columns("PackingID").Text = IIf(CmbCPackName.ListIndex > 0, CmbCPackName.ItemData(CmbCPackName.ListIndex), "")
                  ClaimGrid.Columns("Pack").Value = IIf(Val(TxtCMultiplier.Text) = 0, 0, Val(TxtCMultiplier.Text))
                  ClaimGrid.Columns("EQtyPack").Value = IIf(Val(TxtCEQtyPack.Text) = 0, 0, Val(TxtCEQtyPack.Text))
                  ClaimGrid.Columns("EQtyLoose").Value = Val(TxtCEQtyLoose.Text)
                  ClaimGrid.Columns("DQtyPack").Value = IIf(Val(TxtCDQtyPack.Text) = 0, 0, Val(TxtCDQtyPack.Text))
                  ClaimGrid.Columns("DQtyLoose").Value = Val(TxtCDQtyLoose.Text)
                  ClaimGrid.Columns("Cost").Value = Val(TxtCCost.Text)
                  ClaimGrid.Columns("Amount").Value = Val(TxtCAmount.Text)
                  RsBodyClaim!PackingID = IIf(CmbCPackName.ListIndex > 0, CmbCPackName.ItemData(CmbCPackName.ListIndex), Null)
                  RsBodyClaim!Multiplier = Val(TxtCMultiplier.Text)
                  RsBodyClaim!EQtyPack = Val(TxtCEQtyPack.Text)
                  RsBodyClaim!EQtyLoose = Val(TxtCEQtyLoose.Text)
                  RsBodyClaim!DQtyPack = Val(TxtCDQtyPack.Text)
                  RsBodyClaim!DQtyLoose = Val(TxtCDQtyLoose.Text)
                  RsBodyClaim!Cost = Val(TxtCCost.Text)
                  RsBodyClaim!Amount = Val(TxtCAmount.Text)
                  ClaimGrid.MoveLast
                  Call SubClearDetailAreaClaim
                  TxtCCode.SetFocus
                  ClaimGrid.Redraw = True
                  Exit Sub
               End If
               ClaimGrid.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         SubClearDetailAreaClaim
         ClaimGrid.MoveLast
         TxtCCode.SetFocus
         Exit Sub
      End If
   End If
   ClaimGrid.Redraw = False
   With ClaimGrid
      If TxtCCode.Enabled = True Then
         TxtClaimTotalAmount.Text = Val(TxtClaimTotalAmount.Text) + Val(TxtCAmount.Text)
      Else
         TxtClaimTotalAmount.Text = Val(TxtClaimTotalAmount.Text) + Val(TxtCAmount.Text) - Val(.Columns("Amount").Text)
      End If
      .Columns("ProductName").Text = TxtCProductName.Text
      .Columns("PackName").Text = CmbCPackName.Text
      .Columns("PackingID").Text = IIf(CmbCPackName.ListIndex > 0, CmbCPackName.ItemData(CmbCPackName.ListIndex), "")
      .Columns("Pack").Value = IIf(Val(TxtCMultiplier.Text) = 0, 0, Val(TxtCMultiplier.Text))
      .Columns("EQtyPack").Value = IIf(Val(TxtCEQtyPack.Text) = 0, 0, Val(TxtCEQtyPack.Text))
      .Columns("EQtyLoose").Value = Val(TxtCEQtyLoose.Text)
      .Columns("DQtyPack").Value = IIf(Val(TxtCDQtyPack.Text) = 0, 0, Val(TxtCDQtyPack.Text))
      .Columns("DQtyLoose").Value = Val(TxtCDQtyLoose.Text)
      .Columns("Cost").Value = Val(TxtCCost.Text)
      .Columns("Amount").Value = Val(TxtCAmount.Text)
      RsBodyClaim!PackingID = IIf(CmbCPackName.ListIndex > 0, CmbCPackName.ItemData(CmbCPackName.ListIndex), Null)
      RsBodyClaim!Multiplier = Val(TxtCMultiplier.Text)
      RsBodyClaim!EQtyPack = Val(TxtCEQtyPack.Text)
      RsBodyClaim!EQtyLoose = Val(TxtCEQtyLoose.Text)
      RsBodyClaim!DQtyPack = Val(TxtCDQtyPack.Text)
      RsBodyClaim!DQtyLoose = Val(TxtCDQtyLoose.Text)
      RsBodyClaim!Cost = Val(TxtCCost.Text)
      RsBodyClaim!Amount = Val(TxtCAmount.Text)
      .MoveLast
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   Call SubClearDetailAreaClaim
   TxtCCode.SetFocus
   ClaimGrid.Redraw = True
   Exit Sub
ErrorHandler:
   ClaimGrid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToReplyGrid()
   Dim vrowcounter As Integer
   If Trim(TxtRCode.Text) = "" Then
      TxtRCode.SetFocus
      Exit Sub
   End If
   If CmbRPackName.ListIndex > 0 Then
      If Trim(TxtRMultiplier.Text) = 0 Then
         TxtRMultiplier.SetFocus
         Exit Sub
      End If
   End If
   If Trim(TxtRQtyPack.Text) = "" And Trim(TxtRQtyLoose.Text) = "" Then
      TxtRQtyPack.SetFocus
      Exit Sub
   End If
On Error GoTo ErrorHandler
   RsBodyReply.Filter = "Code='" & TxtRCode.Text & "'"
   If TxtRCode.Enabled Then
      If RsBodyReply.RecordCount = 0 Then
         RsBodyReply.AddNew
         ReplyGrid.Columns("ProductID").Text = TxtProductID.Text
         ReplyGrid.Columns("Code").Text = TxtRCode.Text
         RsBodyReply!Productid = TxtProductID.Text
         RsBodyReply!Code = TxtRCode.Text
      Else
         ReplyGrid.Redraw = False
         ReplyGrid.MoveFirst
            For vrowcounter = 1 To ReplyGrid.Rows
               If ReplyGrid.Columns("Code").Text = TxtProductID.Text Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailAreaReply
                  TxtRQtyLoose.Text = Val(TxtRQtyLoose.Text) + ReplyGrid.Columns("QtyLoose").Value
                  TxtRQtyPack.Text = Val(TxtRQtyPack.Text) + ReplyGrid.Columns("QtyPack").Value
                  TxtReplyTotalAmount.Text = Val(TxtReplyTotalAmount.Text) + Val(TxtRAmount.Text) - Val(ReplyGrid.Columns("Amount").Text)
                  ReplyGrid.Columns("ProductName").Text = TxtRProductName.Text
                  ReplyGrid.Columns("PackingID").Text = IIf(CmbRPackName.ListIndex > 0, CmbRPackName.ItemData(CmbRPackName.ListIndex), "")
                  ReplyGrid.Columns("PackName").Text = CmbRPackName.Text
                  ReplyGrid.Columns("Pack").Value = IIf(Val(TxtRMultiplier.Text) = 0, 0, Val(TxtRMultiplier.Text))
                  ReplyGrid.Columns("RQtyPack").Value = IIf(Val(TxtRQtyPack.Text) = 0, 0, Val(TxtRQtyPack.Text))
                  ReplyGrid.Columns("RQtyLoose").Value = Val(TxtRQtyLoose.Text)
                  ReplyGrid.Columns("Price").Value = Val(TxtRCode.Text)
                  ReplyGrid.Columns("Amount").Value = Val(TxtRAmount.Text)
                  RsBodyReply!PackingID = IIf(CmbRPackName.ListIndex > 0, CmbRPackName.ItemData(CmbRPackName.ListIndex), Null)
                  RsBodyReply!Multiplier = Val(TxtRMultiplier.Text)
                  RsBodyReply!QtyPack = Val(TxtRQtyPack.Text)
                  RsBodyReply!QtyLoose = Val(TxtRQtyLoose.Text)
                  RsBodyReply!Cost = Val(TxtRPrice.Text)
                  RsBodyReply!Amount = Val(TxtRAmount.Text)
                  ReplyGrid.MoveLast
                  Call SubClearDetailAreaReply
                  TxtRCode.SetFocus
                  ReplyGrid.Redraw = True
                  Exit Sub
               End If
               ReplyGrid.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         SubClearDetailAreaReply
         ReplyGrid.MoveLast
         TxtRCode.SetFocus
         Exit Sub
      End If
   End If
   ReplyGrid.Redraw = False
   With ReplyGrid
      If TxtRCode.Enabled = True Then
         TxtReplyTotalAmount.Text = Val(TxtReplyTotalAmount.Text) + Val(TxtRAmount.Text)
      Else
         TxtReplyTotalAmount.Text = Val(TxtReplyTotalAmount.Text) + Val(TxtRAmount.Text) - Val(.Columns("Amount").Text)
      End If
      .Columns("ProductName").Text = TxtRProductName.Text
      .Columns("PackName").Text = CmbRPackName.Text
      .Columns("PackingID").Text = IIf(CmbRPackName.ListIndex > 0, CmbRPackName.ItemData(CmbRPackName.ListIndex), "")
      .Columns("Pack").Value = IIf(Val(TxtRMultiplier.Text) = 0, 0, Val(TxtRMultiplier.Text))
      .Columns("RQtyPack").Value = IIf(Val(TxtRQtyPack.Text) = 0, 0, Val(TxtRQtyPack.Text))
      .Columns("RQtyLoose").Value = Val(TxtRQtyLoose.Text)
      .Columns("Price").Value = Val(TxtRPrice.Text)
      .Columns("Amount").Value = Val(TxtRAmount.Text)
      RsBodyReply!PackingID = IIf(CmbRPackName.ListIndex > 0, CmbRPackName.ItemData(CmbRPackName.ListIndex), Null)
      RsBodyReply!Multiplier = Val(TxtRMultiplier.Text)
      RsBodyReply!QtyPack = Val(TxtRQtyPack.Text)
      RsBodyReply!QtyLoose = Val(TxtRQtyLoose.Text)
      RsBodyReply!Cost = Val(TxtRPrice.Text)
      RsBodyReply!Amount = Val(TxtRAmount.Text)
      .MoveLast
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   Call SubClearDetailAreaReply
   TxtRCode.SetFocus
   ReplyGrid.Redraw = True
   Exit Sub
ErrorHandler:
   ReplyGrid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailAreaReply()
   TxtRCode.Enabled = True
   BtnRProduct.Enabled = True
   TxtRCode.Text = ""
   TxtRProductName.Text = ""
   CmbRPackName.ListIndex = 0
   TxtRMultiplier.Text = ""
   TxtRQtyPack.Text = ""
   TxtRQtyLoose.Text = ""
   TxtRPrice.Text = ""
   TxtRAmount.Text = ""
End Sub

Private Sub SubClearDetailAreaClaim()
   TxtCCode.Enabled = True
   BtnCProduct.Enabled = True
   TxtCCode.Text = ""
   TxtCProductName.Text = ""
   CmbCPackName.ListIndex = 0
   TxtCMultiplier.Text = ""
   TxtCEQtyPack.Text = ""
   TxtCEQtyLoose.Text = ""
   TxtCDQtyPack.Text = ""
   TxtCDQtyLoose.Text = ""
   TxtCCost.Text = ""
   TxtCAmount.Text = ""
End Sub

Private Sub GetDataBackFromClaimGridToTexBoxes()
   On Error GoTo ErrorHandler
   With ClaimGrid
      TxtProductID.Text = .Columns("ProductID").Text
      TxtCCode.Text = .Columns("Code").Text
      TxtCProductName.Text = .Columns("ProductName").Text
      If .Columns("PackName").Text = "" Then
         CmbCPackName.ListIndex = 0
      Else
         CmbCPackName.Text = .Columns("PackName").Text
      End If
      TxtCMultiplier.Text = .Columns("Pack").Text
      TxtCEQtyLoose.Text = .Columns("EQtyLoose").Text
      TxtCEQtyPack.Text = .Columns("EQtyPack").Text
      TxtCDQtyLoose.Text = .Columns("DQtyLoose").Text
      TxtCDQtyPack.Text = .Columns("DQtyPack").Text
      TxtCCost.Text = .Columns("Cost").Text
      TxtCAmount.Text = .Columns("Amount").Value
   End With
         vStrSQL = "select isnull(dbo.FunStock('" & TxtCCode.Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpClaimDate.DateValue + 1 & "',0),0)"
         vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtCCode.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbCPackName.Text
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtCCode.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
         
   If ClaimGrid.Rows = 1 Then ClaimGrid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataBackFromReplyGridToTexBoxes()
   On Error GoTo ErrorHandler
   With ReplyGrid
      TxtProductID.Text = .Columns("ProductID").Text
      TxtRCode.Text = .Columns("Code").Text
      TxtRProductName.Text = .Columns("ProductName").Text
      If .Columns("PackName").Text = "" Then
         CmbRPackName.ListIndex = 0
      Else
         CmbRPackName.Text = .Columns("PackName").Text
      End If
      TxtRMultiplier.Text = .Columns("Pack").Text
      TxtRQtyLoose.Text = .Columns("RQtyLoose").Text
      TxtRQtyPack.Text = .Columns("RQtyPack").Text
      TxtRPrice.Text = .Columns("Price").Text
      TxtRAmount.Text = .Columns("Amount").Value
   End With
         vStrSQL = "select isnull(dbo.FunStock('" & TxtRCode.Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpReplyDate.DateValue + 1 & "',0),0)"
         vQtyLoose = cn.Execute(vStrSQL).Fields(0).Value
         LblStock.Caption = cn.Execute("SELECT dbo.FunGetPack('" & TxtRCode.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbCPackName.Text
         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtRCode.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
   If ReplyGrid.Rows = 1 Then ReplyGrid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetClaim()
   On Error GoTo ErrorHandler
   ssql = "select * FROM ExpiryClaimsHeader where ClaimID=" & Val(TxtClaimID.Text) & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID)
   With cn.Execute(ssql)
      If Not .BOF Then
         DtpClaimDate.DateValue = !ClaimDate
'         TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
'         TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
         'TxtTotalAmount.Text = !TotalAmount
      End If
      .Close
   End With
   ssql = "Select h.*, StoreName FROM ExpiryReplyHeader h inner join stores s on s.storeid = h.storeid where ReplyID=" & Val(TxtClaimID.Text)
   With cn.Execute(ssql)
      If Not .BOF Then
         DtpReplyDate.DateValue = !ReplyDate
         TxtRepliedAmount.Text = IIf(IsNull(!RepliedAmount), "", !RepliedAmount)
         TxtStoreID.Text = !StoreID
         TxtStoreName.Text = !StoreName
         
         'TxtTotalAmount.Text = !TotalAmount
      End If
      .Close
   End With
   Call PopulateDataToClaim
   Call PopulateDataToReply
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   ClaimGrid.Redraw = True
   ReplyGrid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtCDQtyLoose_LostFocus()
   Call GetDataFromTexBoxesToClaimGrid
End Sub

Private Sub TxtCEQtyLoose_Change()
   Call SubCalculateBodyClaim
End Sub

Private Sub TxtCEQtyPack_Change()
   Call SubCalculateBodyClaim
End Sub

Private Sub TxtCMultiplier_Change()
   Call SubCalculateBodyClaim
End Sub

Private Sub TxtCDQtyLoose_Change()
   Call SubCalculateBodyClaim
End Sub

Private Sub TxtCDQtyPack_Change()
   Call SubCalculateBodyClaim
End Sub

Private Sub TxtCCode_Change()
   If ActiveControl.Name <> TxtCCode.Name Then Exit Sub
   If TxtCProductName.Text <> "" Then
      TxtCProductName.Text = ""
      TxtCCost.Text = ""
   End If
End Sub

Private Sub TxtCCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then ClaimGrid.SetFocus
End Sub

Private Sub TxtCCode_Validate(Cancel As Boolean)
   If TxtCProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtCCode.Text) = "" Then Exit Sub
   vTemp = FunSelectClaimProduct(ssValidate, False)
   If vTemp = False Then   '   vTemp = FunSelectClaimProduct(ssButton, False)
      Cancel = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtRCode_Change()
   If ActiveControl.Name <> TxtRCode.Name Then Exit Sub
   If TxtRProductName.Text <> "" Then
      TxtRProductName.Text = ""
      TxtRPrice.Text = ""
   End If
End Sub

Private Sub TxtRCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then ReplyGrid.SetFocus
End Sub

Private Sub TxtRCode_Validate(Cancel As Boolean)
   If TxtRProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtRCode.Text) = "" Then Exit Sub
   vTemp = FunSelectReplyProduct(ssValidate, False)
   If vTemp = False Then   '   vTemp = FunSelectProduct(ssButton, False)
      Cancel = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtRPrice_LostFocus()
   Call GetDataFromTexBoxesToReplyGrid
End Sub

Private Sub TxtRMultiplier_Change()
   Call SubCalculateBodyReply
End Sub

Private Sub TxtRQtyLoose_Change()
   Call SubCalculateBodyReply
End Sub

Private Sub TxtRQtyPack_Change()
   Call SubCalculateBodyReply
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

Private Function FunRefreshData() As Boolean
On Error GoTo ErrorHandler
  Dim vSQL As String
  Me.MousePointer = vbHourglass
  vSQL = "Select H.*,B.productID, B.Code,B.packingID,B.Multiplier, B.EQtyPack, B.EQtyLoose, B.Cost, B.Amount, B.DQtyPack, B.DQtyLoose," & vbCrLf _
                + " ProductName, PackingName from ExpiryClaimsHeader H" & vbCrLf _
                + " Inner Join ExpiryClaimsBody B on H.ClaimID = B.claimID" & vbCrLf _
                + " Inner Join Products P on P.productID = b.Productid" & vbCrLf _
                + " Left Outer Join Packings Pack on Pack.PackingID = B.PackingID Where H.ClaimID =" & Val(TxtClaimID.Text)
    Set RsBodyClaim = cn.Execute(vSQL)
    
  vSQL = "Select H.*,B.productID, B.Code,B.packingID,B.Multiplier, B.QtyPack, B.QtyLoose, B.Cost, B.Amount, ProductName, PackingName, StoreName from ExpiryReplyHeader H " & vbCrLf _
              + " Inner Join ExpiryReplyBody B on H.ReplyID = B.ReplyID" & vbCrLf _
              + " Inner Join Products P on P.productID = b.Productid" & vbCrLf _
              + " Inner Join Stores  S on S.StoreID = H.StoreID" & vbCrLf _
              + " Left Outer Join Packings Pack on Pack.PackingID = B.PackingID Where H.ReplyID = " & TxtClaimID.Text
  Set RsBodyReply = cn.Execute(vSQL)
  
  FunRefreshData = True
  Me.MousePointer = vbDefault
  Exit Function
ErrorHandler:
  Me.MousePointer = vbDefault
  FunRefreshData = False
  Call ShowErrorMessage
End Function

Private Sub SetCrystalReport1()
  On Error GoTo ErrorHandler
  Me.MousePointer = vbHourglass
  Set RptReportViewer.Report = New CrptExpiryClaimInv
  'this code works through the RDC object model to identify a subreport object
  'in the main report
Dim crSecs As CRAXDRT.Sections
Dim crSec As CRAXDRT.Section
Dim crRepObjs As CRAXDRT.ReportObjects
Dim crSubRepObj As CRAXDRT.SubreportObject
Dim crSubReport As CRAXDRT.Report
Dim i As Integer
Dim x As Integer
Dim Y As Integer
Set crSecs = RptReportViewer.Report.Sections
For i = 1 To crSecs.Count
  Set crSec = crSecs.Item(i)
  Set crRepObjs = crSec.ReportObjects
    For x = 1 To crRepObjs.Count
      If crRepObjs.Item(x).Kind = crSubreportObject Then
            Set crSubReport = RptReportViewer.Report.OpenSubreport(crRepObjs.Item(x).SubreportName)
            'the following code sets the subreport table to a different database
            crSubReport.Database.SetDataSource RsBodyReply, 3, 1
            For Y = 1 To crSubReport.Sections.Count
               crSubReport.Sections(Y).Suppress = True
            Next Y
            'set the value for a text object in the header of the subreport
            'CRReport.Subreport1_Text2.SetText "This is the subreport"
            'within this loop you can set other properties of the subreport and
            'the field objects and sections in it.
      End If
    Next x
Next i
  'RptReportViewer.Report.TxtCompanyName.SetText ObjSupernetRegistry.CompanyName
  RptReportViewer.Report.ReportTitle = "Expiry / Damage Claim Invoice"
  'RptReportViewer.Report.ParameterFields(1).AddCurrentValue "" 'Format(DtpTo.Value, "dd/MM/yyyy")
  RptReportViewer.Report.Database.SetDataSource RsBodyClaim, 3, 1
  RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
  RptReportViewer.Report.PaperOrientation = crPortrait
  Me.MousePointer = vbDefault
  Exit Sub
ErrorHandler:
  Me.MousePointer = vbDefault
  Call ShowErrorMessage
End Sub

Private Sub SetCrystalReport()
  On Error GoTo ErrorHandler
  Me.MousePointer = vbHourglass
  Set RptReportViewer.Report = New CrptExpiryClaimInv
  'this code works through the RDC object model to identify a subreport object
  'in the main report
Dim crSecs As CRAXDRT.Sections
Dim crSec As CRAXDRT.Section
Dim crRepObjs As CRAXDRT.ReportObjects
Dim crSubRepObj As CRAXDRT.SubreportObject
Dim crSubReport As CRAXDRT.Report
Dim i As Integer
Dim x As Integer
Set crSecs = RptReportViewer.Report.Sections
For i = 1 To crSecs.Count
  Set crSec = crSecs.Item(i)
  Set crRepObjs = crSec.ReportObjects
    For x = 1 To crRepObjs.Count
      If crRepObjs.Item(x).Kind = crSubreportObject Then
         'If X = 1 And i = 4 Then
            Set crSubReport = RptReportViewer.Report.OpenSubreport(crRepObjs.Item(x).SubreportName)
            'the following code sets the subreport table to a different database
            crSubReport.Database.SetDataSource RsBodyReply, 3, 1
            'set the value for a text object in the header of the subreport
            'CRReport.Subreport1_Text2.SetText "This is the subreport"
            'within this loop you can set other properties of the subreport and
            'the field objects and sections in it.
         'ElseIf X = 1 And i = 5 Then
         '   Set crSubReport = RptReportViewer.Report.OpenSubreport(crRepObjs.Item(X).SubreportName)
         '   crSubReport.Database.SetDataSource Rs1, 3, 1
         'End If
      End If
    Next
Next
  'RptReportViewer.Report.TxtCompanyName.SetText ObjSupernetRegistry.CompanyName
  RptReportViewer.Report.ReportTitle = "Expiry / Damage Claim Invoice"
  'RptReportViewer.Report.ParameterFields(1).AddCurrentValue "" 'Format(DtpTo.Value, "dd/MM/yyyy")
  RptReportViewer.Report.Database.SetDataSource RsBodyClaim, 3, 1
  RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
  RptReportViewer.Report.PaperOrientation = crPortrait
  Me.MousePointer = vbDefault
  Exit Sub
ErrorHandler:
  Me.MousePointer = vbDefault
  Call ShowErrorMessage
End Sub

Private Sub SetReportParameterField()
On Error GoTo ErrorHandler
    With cn.Execute("Select CompanyName,Address,City,PhoneNo,email from Company")
      If .RecordCount > 0 Then
         RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(IsNull(!CompanyName), "", CStr(!CompanyName))
         RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(IsNull(!Address), "", !Address) & IIf(IsNull(!City), "", ", " & !City & ".")
         RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(IsNull(!PhoneNo), "", CStr(!PhoneNo))
      End If
    .Close
    End With
    RptReportViewer.Report.ParameterFields(4).AddCurrentValue cn.Execute("Select Name from Manufacturer").Fields(0).Value
Exit Sub
ErrorHandler:
  Me.MousePointer = vbDefault
  Call ShowErrorMessage
End Sub

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
'      TxtCustomerID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub
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

