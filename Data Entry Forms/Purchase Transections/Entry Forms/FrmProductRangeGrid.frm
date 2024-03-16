VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmProductRangeGrid 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmProductRangeGrid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmShowQty 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   720
      TabIndex        =   55
      Top             =   1215
      Width           =   3735
      Begin VB.OptionButton RdoNone 
         BackColor       =   &H00FFFFFF&
         Caption         =   "None"
         Height          =   255
         Left            =   135
         TabIndex        =   58
         Top             =   0
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.OptionButton RdoSaleQty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Sale Qty"
         Height          =   255
         Left            =   135
         TabIndex        =   57
         Top             =   240
         Width           =   1725
      End
      Begin VB.OptionButton RdoStockSale 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Only Negative When Stock Minus Sale"
         Height          =   255
         Left            =   135
         TabIndex        =   56
         Top             =   495
         Width           =   3570
      End
   End
   Begin VB.Frame FrmProductPrices 
      Height          =   1095
      Left            =   9045
      TabIndex        =   48
      Top             =   9225
      Visible         =   0   'False
      Width           =   6270
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridProductPrices 
         Height          =   885
         Left            =   60
         TabIndex        =   49
         Top             =   150
         Width           =   6135
         ScrollBars      =   0
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   5
         stylesets.count =   3
         stylesets(0).Name=   "SelectedCol"
         stylesets(0).ForeColor=   0
         stylesets(0).BackColor=   12713983
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "FrmProductRangeGrid.frx":0ECA
         stylesets(1).Name=   "Select"
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
         stylesets(1).Picture=   "FrmProductRangeGrid.frx":0EE6
         stylesets(2).Name=   "SelectedRow"
         stylesets(2).ForeColor=   16777215
         stylesets(2).BackColor=   8388608
         stylesets(2).HasFont=   -1  'True
         BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(2).Picture=   "FrmProductRangeGrid.frx":0F02
         MultiLine       =   0   'False
         ActiveCellStyleSet=   "SelectedCol"
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
         SelectTypeRow   =   0
         ForeColorEven   =   0
         ForeColorOdd    =   8388736
         BackColorEven   =   16776960
         RowHeight       =   714
         Columns.Count   =   5
         Columns(0).Width=   1402
         Columns(0).Caption=   "Pur"
         Columns(0).Name =   "Pur"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   5
         Columns(0).FieldLen=   256
         Columns(1).Width=   1402
         Columns(1).Caption=   "List"
         Columns(1).Name =   "List"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   5
         Columns(1).FieldLen=   256
         Columns(2).Width=   1402
         Columns(2).Caption=   "WS"
         Columns(2).Name =   "WS"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   5
         Columns(2).FieldLen=   256
         Columns(3).Width=   1402
         Columns(3).Caption=   "Retail"
         Columns(3).Name =   "Retail"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   5
         Columns(3).FieldLen=   256
         Columns(4).Width=   5239
         Columns(4).Caption=   "Description"
         Columns(4).Name =   "Description"
         Columns(4).CaptionAlignment=   2
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(4).Locked=   -1  'True
         TabNavigation   =   1
         _ExtentX        =   10821
         _ExtentY        =   1561
         _StockProps     =   79
         Caption         =   "Product Prices"
         ForeColor       =   0
         BackColor       =   16776960
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame FrmHistory 
      Height          =   1275
      Left            =   360
      TabIndex        =   46
      Top             =   9225
      Visible         =   0   'False
      Width           =   3915
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridHistory 
         Height          =   1050
         Left            =   45
         TabIndex        =   47
         Top             =   135
         Width           =   3780
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   4
         stylesets.count =   3
         stylesets(0).Name=   "SelectedCol"
         stylesets(0).ForeColor=   0
         stylesets(0).BackColor=   12713983
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "FrmProductRangeGrid.frx":0F1E
         stylesets(1).Name=   "Select"
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
         stylesets(1).Picture=   "FrmProductRangeGrid.frx":0F3A
         stylesets(2).Name=   "SelectedRow"
         stylesets(2).ForeColor=   16777215
         stylesets(2).BackColor=   8388608
         stylesets(2).HasFont=   -1  'True
         BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(2).Picture=   "FrmProductRangeGrid.frx":0F56
         MultiLine       =   0   'False
         ActiveCellStyleSet=   "SelectedCol"
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
         SelectTypeRow   =   0
         ForeColorEven   =   0
         BackColorOdd    =   15724527
         RowHeight       =   423
         Columns.Count   =   4
         Columns(0).Width=   1588
         Columns(0).Caption=   "PurPrice"
         Columns(0).Name =   "PurPrice"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1746
         Columns(1).Caption=   "WS Price"
         Columns(1).Name =   "WSPrice"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1244
         Columns(2).Caption=   "Margin"
         Columns(2).Name =   "Margin"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1614
         Columns(3).Caption=   "Margin Per"
         Columns(3).Name =   "MarginPer"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   6667
         _ExtentY        =   1852
         _StockProps     =   79
         Caption         =   "History"
         BackColor       =   15724527
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CheckBox ChkSaleDataStock 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Check Sale Data Stock"
      Height          =   255
      Left            =   720
      TabIndex        =   45
      Top             =   945
      Width           =   2010
   End
   Begin VB.ComboBox CmbBrand 
      Height          =   315
      Left            =   10620
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   10770
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.ComboBox CmbSubGroup 
      Height          =   315
      Left            =   8955
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   10770
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.TextBox TxtToProductID 
      Height          =   345
      Left            =   2460
      TabIndex        =   2
      Top             =   2280
      Width           =   1380
   End
   Begin VB.ComboBox CmbSortBy 
      Height          =   315
      Left            =   11760
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1725
      Width           =   1170
   End
   Begin VB.TextBox TxtFromProductID 
      Height          =   345
      Left            =   690
      TabIndex        =   1
      Top             =   2280
      Width           =   1380
   End
   Begin VB.TextBox TxtProductName 
      Height          =   345
      Left            =   690
      TabIndex        =   3
      Top             =   2910
      Width           =   2700
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   5370
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   10770
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   7275
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   10770
      Visible         =   0   'False
      Width           =   1665
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
      Height          =   1635
      Left            =   13200
      TabIndex        =   13
      Top             =   540
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
         Height          =   1545
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Tag             =   "NC"
         Text            =   "FrmProductRangeGrid.frx":0F72
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
         TabIndex        =   15
         Top             =   90
         Width           =   135
      End
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   5865
      TabIndex        =   9
      Top             =   8715
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Select"
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
      MICON           =   "FrmProductRangeGrid.frx":0FC4
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7170
      TabIndex        =   10
      Top             =   8715
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
      MICON           =   "FrmProductRangeGrid.frx":0FE0
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8460
      TabIndex        =   11
      Top             =   8715
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
      MICON           =   "FrmProductRangeGrid.frx":0FFC
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnFilter 
      Height          =   315
      Left            =   3465
      TabIndex        =   17
      Top             =   2895
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      TX              =   "Filter"
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
      MICON           =   "FrmProductRangeGrid.frx":1018
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnApply 
      Height          =   315
      Left            =   10350
      TabIndex        =   18
      Top             =   1770
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      TX              =   "Apply All"
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
      MICON           =   "FrmProductRangeGrid.frx":1034
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtQtyLoose 
      Height          =   315
      Left            =   9330
      TabIndex        =   19
      Top             =   1770
      Width           =   795
      _ExtentX        =   1402
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5190
      Left            =   90
      TabIndex        =   28
      Top             =   3420
      Width           =   15180
      ScrollBars      =   3
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   27
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
      stylesets(0).Picture=   "FrmProductRangeGrid.frx":1050
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
      stylesets(1).Picture=   "FrmProductRangeGrid.frx":106C
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedCol"
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
      SelectTypeRow   =   0
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   106
      Columns.Count   =   27
      Columns(0).Width=   1138
      Columns(0).Caption=   "P ID"
      Columns(0).Name =   "ID"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   6324
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "Branch 1"
      Columns(2).Name =   "Branch1"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "Branch 2"
      Columns(3).Name =   "Branch2"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "Branch 3"
      Columns(4).Name =   "Branch3"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "Branch 4"
      Columns(5).Name =   "Branch4"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "Branch 5"
      Columns(6).Name =   "Branch5"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "Branch 6"
      Columns(7).Name =   "Branch6"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "Branch 7"
      Columns(8).Name =   "Branch7"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "Branch 8"
      Columns(9).Name =   "Branch8"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "Branch 9"
      Columns(10).Name=   "Branch9"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   1376
      Columns(11).Caption=   "Stock"
      Columns(11).Name=   "Stock"
      Columns(11).CaptionAlignment=   2
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(11).Locked=   -1  'True
      Columns(12).Width=   1376
      Columns(12).Caption=   "QtyPack"
      Columns(12).Name=   "QtyPack"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   1085
      Columns(13).Caption=   "Qty (L)"
      Columns(13).Name=   "QtyLoose"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   1376
      Columns(14).Caption=   "Pur Price"
      Columns(14).Name=   "PurPrice"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   5
      Columns(14).FieldLen=   256
      Columns(15).Width=   1455
      Columns(15).Caption=   "List Price"
      Columns(15).Name=   "ListPrice"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(15).Locked=   -1  'True
      Columns(16).Width=   1482
      Columns(16).Caption=   "WS Price"
      Columns(16).Name=   "WSPrice"
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   8
      Columns(16).FieldLen=   256
      Columns(16).Locked=   -1  'True
      Columns(17).Width=   1746
      Columns(17).Caption=   "Retail Price"
      Columns(17).Name=   "RetailPrice"
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   8
      Columns(17).FieldLen=   256
      Columns(18).Width=   1191
      Columns(18).Caption=   "Disc %"
      Columns(18).Name=   "DiscPer"
      Columns(18).DataField=   "Column 18"
      Columns(18).DataType=   8
      Columns(18).FieldLen=   256
      Columns(19).Width=   1217
      Columns(19).Caption=   "DiscPC"
      Columns(19).Name=   "DiscPC"
      Columns(19).DataField=   "Column 19"
      Columns(19).DataType=   8
      Columns(19).FieldLen=   256
      Columns(20).Width=   2223
      Columns(20).Caption=   "Packing Name"
      Columns(20).Name=   "PackingName"
      Columns(20).DataField=   "Column 20"
      Columns(20).DataType=   8
      Columns(20).FieldLen=   256
      Columns(20).Locked=   -1  'True
      Columns(21).Width=   1376
      Columns(21).Caption=   "Multiplier"
      Columns(21).Name=   "multiplier"
      Columns(21).DataField=   "Column 21"
      Columns(21).DataType=   8
      Columns(21).FieldLen=   256
      Columns(21).Locked=   -1  'True
      Columns(22).Width=   2461
      Columns(22).Caption=   "Company"
      Columns(22).Name=   "Company"
      Columns(22).CaptionAlignment=   2
      Columns(22).DataField=   "Column 22"
      Columns(22).DataType=   8
      Columns(22).NumberFormat=   "########.##"
      Columns(22).FieldLen=   256
      Columns(22).Locked=   -1  'True
      Columns(23).Width=   2461
      Columns(23).Caption=   "Group"
      Columns(23).Name=   "Group"
      Columns(23).CaptionAlignment=   2
      Columns(23).DataField=   "Column 23"
      Columns(23).DataType=   8
      Columns(23).FieldLen=   256
      Columns(23).Locked=   -1  'True
      Columns(24).Width=   2461
      Columns(24).Caption=   "Sub Group"
      Columns(24).Name=   "SubGroup"
      Columns(24).CaptionAlignment=   2
      Columns(24).DataField=   "Column 24"
      Columns(24).DataType=   8
      Columns(24).FieldLen=   256
      Columns(24).Locked=   -1  'True
      Columns(25).Width=   2461
      Columns(25).Caption=   "Brand"
      Columns(25).Name=   "Brand"
      Columns(25).CaptionAlignment=   2
      Columns(25).DataField=   "Column 25"
      Columns(25).DataType=   8
      Columns(25).FieldLen=   256
      Columns(26).Width=   3200
      Columns(26).Caption=   "GroupID"
      Columns(26).Name=   "GroupID"
      Columns(26).DataField=   "Column 26"
      Columns(26).DataType=   8
      Columns(26).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   26776
      _ExtentY        =   9155
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
   Begin JeweledBut.JeweledButton BtnFromProduct 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2100
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2280
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
      MICON           =   "FrmProductRangeGrid.frx":1088
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnToProduct 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3870
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2265
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
      MICON           =   "FrmProductRangeGrid.frx":10A4
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPurPrice 
      Height          =   315
      Left            =   8415
      TabIndex        =   31
      Top             =   1770
      Width           =   795
      _ExtentX        =   1402
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
   Begin SITextBox.Txt TxtDiscPer 
      Height          =   315
      Left            =   7470
      TabIndex        =   33
      Top             =   1770
      Width           =   795
      _ExtentX        =   1402
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
   Begin JeweledBut.JeweledButton BtnVender 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9390
      TabIndex        =   35
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1155
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
      MICON           =   "FrmProductRangeGrid.frx":10C0
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtVenderName 
      Height          =   315
      Left            =   9750
      TabIndex        =   36
      Tag             =   "nc"
      Top             =   1155
      Width           =   3135
      _ExtentX        =   5530
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
   Begin SITextBox.Txt TxtVenderID 
      Height          =   315
      Left            =   8370
      TabIndex        =   0
      Top             =   1155
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   5
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
   Begin JeweledBut.JeweledButton BtnUpdateStock 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5760
      TabIndex        =   39
      Top             =   1725
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      TX              =   "Update Stock"
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
      MICON           =   "FrmProductRangeGrid.frx":10DC
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   5640
      TabIndex        =   40
      Top             =   1155
      Visible         =   0   'False
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
      Left            =   6990
      TabIndex        =   41
      Top             =   1155
      Visible         =   0   'False
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
   Begin JeweledBut.JeweledButton BtnCompany 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6750
      TabIndex        =   50
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2385
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
      MICON           =   "FrmProductRangeGrid.frx":10F8
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   5745
      TabIndex        =   51
      Top             =   2385
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
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
      IntegralPoint   =   9
   End
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   10965
      TabIndex        =   59
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2385
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
      MICON           =   "FrmProductRangeGrid.frx":1114
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   9945
      TabIndex        =   53
      Top             =   2385
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
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
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtCompanyName 
      Height          =   315
      Left            =   7110
      TabIndex        =   62
      Top             =   2385
      Width           =   2265
      _ExtentX        =   3995
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
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   11340
      TabIndex        =   63
      Tag             =   "nc"
      Top             =   2385
      Width           =   2265
      _ExtentX        =   3995
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
   Begin JeweledBut.JeweledButton BtnSubGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   10965
      TabIndex        =   64
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2985
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
      MICON           =   "FrmProductRangeGrid.frx":1130
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubGroupID 
      Height          =   315
      Left            =   9945
      TabIndex        =   54
      Top             =   2985
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
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
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtSubGroupName 
      Height          =   315
      Left            =   11325
      TabIndex        =   67
      Tag             =   "nc"
      Top             =   2985
      Width           =   2265
      _ExtentX        =   3995
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
   Begin JeweledBut.JeweledButton BtnBrand 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6735
      TabIndex        =   68
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2985
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
      MICON           =   "FrmProductRangeGrid.frx":114C
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBrandID 
      Height          =   315
      Left            =   5715
      TabIndex        =   52
      Top             =   2985
      Width           =   1020
      _ExtentX        =   1799
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
      IntegralPoint   =   15
   End
   Begin SITextBox.Txt TxtBrandName 
      Height          =   315
      Left            =   7095
      TabIndex        =   69
      Tag             =   "nc"
      Top             =   2985
      Width           =   2265
      _ExtentX        =   3995
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
   Begin JeweledBut.JeweledButton BtnApplyStock 
      Height          =   450
      Left            =   4635
      TabIndex        =   74
      Top             =   2205
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   794
      TX              =   "Apply Stock"
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
      MICON           =   "FrmProductRangeGrid.frx":1168
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
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
      Left            =   7095
      TabIndex        =   73
      Top             =   2160
      Width           =   1320
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brand Name"
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
      Left            =   7095
      TabIndex        =   72
      Top             =   2790
      Width           =   1050
   End
   Begin VB.Label LblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company ID"
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
      Left            =   5715
      TabIndex        =   71
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brand ID"
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
      Left            =   5715
      TabIndex        =   70
      Top             =   2790
      Width           =   765
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Group Name"
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
      Left            =   11325
      TabIndex        =   66
      Top             =   2790
      Width           =   1455
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Group ID"
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
      Left            =   9945
      TabIndex        =   65
      Top             =   2790
      Width           =   1170
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
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
      Left            =   11325
      TabIndex        =   61
      Top             =   2160
      Width           =   1065
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group ID"
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
      Left            =   9945
      TabIndex        =   60
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Product ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   690
      TabIndex        =   44
      Top             =   2040
      Width           =   1395
   End
   Begin VB.Label LblTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7005
      TabIndex        =   43
      Top             =   960
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label LblFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5640
      TabIndex        =   42
      Top             =   960
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vender ID"
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
      Left            =   8370
      TabIndex        =   38
      Top             =   960
      Width           =   870
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vender Name"
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
      Left            =   9765
      TabIndex        =   37
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7470
      TabIndex        =   34
      Top             =   1545
      Width           =   585
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8370
      TabIndex        =   32
      Top             =   1545
      Width           =   795
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brand Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   10620
      TabIndex        =   27
      Top             =   10545
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubGroup Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8955
      TabIndex        =   26
      Top             =   10545
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Product ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2460
      TabIndex        =   25
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Loose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9285
      TabIndex        =   24
      Top             =   1545
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort BY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   11790
      TabIndex        =   23
      Top             =   1485
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   690
      TabIndex        =   22
      Top             =   2670
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5370
      TabIndex        =   21
      Top             =   10545
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7275
      TabIndex        =   20
      Top             =   10545
      Visible         =   0   'False
      Width           =   1065
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
      Left            =   11070
      TabIndex        =   16
      Top             =   495
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Range Grid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   12
      Top             =   270
      Width           =   2565
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "FrmProductRangeGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vWords() As String
Dim vProductName As String
Dim vProductID As String
Dim Rs As New ADODB.Recordset
Dim vSuppressUpdateEvent As Boolean
Dim vSQL, vSqlBranch, vSqlPivot As String

Public ParaInPartyID As String
Public ParaInBoth As Boolean

Private Function GetGroupID(cmb As ComboBox) As String
    On Error GoTo ErrorHandler
    If cmb.ListIndex < 0 Then Exit Function
    GetGroupID = Chr(Left(cmb.ItemData(cmb.ListIndex), 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 3, 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 5, 2))
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub PopulateGrid()
   On Error GoTo ErrorHandler
   Dim vQtyLoose As Double
   
   Me.MousePointer = vbHourglass
   
   Set RsTemp = New ADODB.Recordset
   RsTemp.Fields.Append "ProductID", adInteger
   RsTemp.Fields.Append "ItemCode", adVarChar, 9
   RsTemp.Fields.Append "ProductName", adVarChar, 200
   RsTemp.Fields.Append "GroupID", adVarChar, 5
   RsTemp.Fields.Append "Price", adDouble
   RsTemp.Fields.Append "QtyPack", adDouble
   RsTemp.Fields.Append "QtyLoose", adDouble
   RsTemp.Fields.Append "PackingName", adVarChar, 100
   RsTemp.Fields.Append "Multiplier", adDouble
   RsTemp.Fields.Append "PartyID", adVarChar, 20
   RsTemp.Fields.Append "DiscPer", adDouble
   RsTemp.Fields.Append "DiscPC", adDouble
   RsTemp.Fields.Append "RetailPrice", adDouble
   
   
   RsTemp.Open

   
   Me.MousePointer = vbHourglass
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
   
   If ObjRegistry.ShowMultiBranches = True Then
      vSqlBranch = " [1] Branch1, [2] Branch2, [3] Branch3, [4] Branch4, [5] Branch5, [6] Branch6, [7] Branch7, [8] Branch8, [9] Branch9,"
      vSqlPivot = " Left Outer Join (" & vbCrLf _
      + " SELECT * FROM (" & vbCrLf _
      + " SELECT ProductID, StoreID, QtyLoose FROM CurrentSTockStore" & vbCrLf _
      + " ) AS SourceTable" & vbCrLf _
      + " PIVOT(" & vbCrLf _
      + " Sum(QtyLoose) FOR [StoreID] IN([1], [2], [3], [4], [5], [6], [7], [8], [9])" & vbCrLf _
      + " ) AS PivotStore )CSS On CSS.ProductID = P.ProductID "
   Else
      vSqlBranch = ""
      vSqlPivot = ""
   End If
   
   
   If TxtVenderID.Text = "" Then
     vSQL = " SELECT P.ProductID, p.itemcode, ProductName, P.GroupID, CS.QtyLoose StockLoose, sb.QtyLoose as QtyLoose, PurPrice, ListPrice, WSPrice, P.Discper, p.DiscPC, GroupName, isnull(CompanyName,'') as CompanyName, " & vbCrLf _
      + vSqlBranch & vbCrLf _
      + " isnull(SubGroupName,'') as SubGroupName, isnull(BrandName,'') as BrandName, RetailPrice," & vbCrLf _
      + " isnull(packingname,'') packingname, isnull(pp.multiplier,0) multiplier, dbo.FunLoseQtyToPackStr(p.ProductID, CS.QtyLoose) Stock" & vbCrLf _
      + " FROM Products p Left Outer join Groups g on g.GroupID = p.GroupID" & vbCrLf _
      + " Left outer join CurrentStock CS on cs.productid = p.productid" & vbCrLf _
      + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID" & vbCrLf _
      + " left outer join Companies c on c.CompanyID = p.CompanyID" & vbCrLf _
      + " Left outer join packings pk on pk.packingid = p.purchasepackingid" & vbCrLf _
      + " Left outer join productpacking pp on pp.productid = p.productid" & vbCrLf _
      + " left outer join Brands b on b.BrandID = p.BrandID" & vbCrLf _
      + vSqlPivot
      vSQL = vSQL + " left outer join (select b.Productid, VendorID, Price as LastPrice, Discper  " & vbCrLf _
            + " from PurchaseHeader h inner join purchasebody b on b.SID = h.SID " & vbCrLf _
            + " inner join (select ProductID, max(SerialNo) SerialNo from purchasebody group by ProductID) d on b.SerialNo = d.SerialNo) Purchase" & vbCrLf _
            + " on Purchase.productid = p.Productid" & vbCrLf _
            + " left outer join (select productid, Sum(Qty + isnull(QtyPack,0)*isnull(multiplier,1)) QtyLoose from SaleBody where BillDate Between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "' Group by Productid ) sb on sb.ProductID = p.ProductID" & vbCrLf _
            + " where p.isLocked = 0 " & vProductID & vProductName
      vSQL = vSQL + "" & vbCrLf _
      + IIf(TxtCompanyID.Text = "", "", " and p.CompanyID in (" & TxtCompanyID.Text & ")") & vbCrLf _
      + IIf(TxtBrandID.Text = "", "", " and p.BrandID in (" & TxtBrandID.Text & ")") & vbCrLf _
      + IIf(TxtGroupID.Text = "", "", " and p.GroupID in (" & TxtGroupID.Text & ")") & vbCrLf _
      + IIf(TxtSubGroupID.Text = "", "", " and p.SubGroupID in (" & TxtSubGroupID.Text & ")") & vbCrLf _
      + IIf(TxtVenderID.Text = "", "", " and isnull(purchase.VendorID,p.Vendorid1) = " & Val(TxtVenderID.Text)) & vbCrLf _
      + IIf(ChkSaleDataStock.Value = 0, "", " and sb.ProductID is not null") & vbCrLf _
      + IIf(RdoStockSale.Value = True, " and cs.qtyloose - sb.qtyloose < 0", "") & vbCrLf _
      + " Order by " + IIf(ObjRegistry.ShowAllPrices, " p.CompanyID,", "") + CmbSortBy.Text
           
   Else
      vSQL = " SELECT P.ProductID, P.ItemCode, ProductName, P.GroupID, CS.QtyLoose StockLoose, sb.QtyLoose as QtyLoose, isnull(LastPrice,p.PurPrice) PurPrice, isnull(Purchase.Discper, case when p.purprice = 0 then 0 else Round(p.purdiscpc * 100 / p.purprice,2) end) Discper, p.purdiscpc as DiscPC , GroupName, isnull(CompanyName,'') as CompanyName, " & vbCrLf _
      + vSqlBranch & vbCrLf _
      + " isnull(SubGroupName,'') as SubGroupName, isnull(BrandName,'') as BrandName, RetailPrice, ListPrice, WSPrice," & vbCrLf _
      + " isnull(packingname,'') packingname, isnull(pp.multiplier,0) multiplier, dbo.FunLoseQtyToPackStr(p.ProductID, CS.QtyLoose) Stock" & vbCrLf _
      + " FROM Products p Left Outer join Groups g on g.GroupID = p.GroupID" & vbCrLf _
      + " Left outer join CurrentStock CS on cs.productid = p.productid" & vbCrLf _
      + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID" & vbCrLf _
      + " left outer join Companies c on c.CompanyID = p.CompanyID" & vbCrLf _
      + " Left outer join packings pk on pk.packingid = p.purchasepackingid" & vbCrLf _
      + " Left outer join productpacking pp on pp.productid = p.productid" & vbCrLf _
      + " left outer join Brands b on b.BrandID = p.BrandID " & vbCrLf _
      + vSqlPivot & vbCrLf _
      + " left outer join (select b.Productid, VendorID, Price as LastPrice, Discper  " & vbCrLf _
      + " from PurchaseHeader h inner join purchasebody b on b.SID = h.SID  " & vbCrLf _
      + " inner join (select ProductID, max(SerialNo) SerialNo from purchasebody group by ProductID) d on b.SerialNo = d.SerialNo) Purchase " & vbCrLf _
      + " on Purchase.productid = p.Productid " & vbCrLf _
      + " left outer join (select productid, Sum(Qty + isnull(QtyPack,0)*isnull(multiplier,1)) QtyLoose from SaleBody where BillDate Between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "' Group by Productid ) sb on sb.ProductID = p.ProductID" & vbCrLf _
      + " where p.isLocked = 0 " & vProductID & vProductName
      vSQL = vSQL + "" & vbCrLf _
      + IIf(TxtCompanyID.Text = "", "", " and p.CompanyID in (" & TxtCompanyID.Text & ")") & vbCrLf _
      + IIf(TxtBrandID.Text = "", "", " and p.BrandID in (" & TxtBrandID.Text & ")") & vbCrLf _
      + IIf(TxtGroupID.Text = "", "", " and p.GroupID in (" & TxtGroupID.Text & ")") & vbCrLf _
      + IIf(TxtSubGroupID.Text = "", "", " and p.SubGroupID in (" & TxtSubGroupID.Text & ")") & vbCrLf _
      + IIf(TxtVenderID.Text = "", "", " and isnull(purchase.VendorID,p.Vendorid1) = " & Val(TxtVenderID.Text)) & vbCrLf _
      + IIf(ChkSaleDataStock.Value = 0, "", " and sb.ProductID is not null") & vbCrLf _
      + IIf(RdoStockSale.Value = True, " and cs.qtyloose - sb.qtyloose < 0", "") & vbCrLf _
      + " Order by " + IIf(ObjRegistry.ShowAllPrices, " p.CompanyID,", "") + CmbSortBy.Text
   End If
   CN.CommandTimeout = 0
   With CN.Execute(vSQL)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        If ObjRegistry.ShowMultiBranches = True Then
         Grid.Columns("Branch1").Value = IIf(IsNull(!Branch1), "", !Branch1)
         Grid.Columns("Branch2").Value = IIf(IsNull(!Branch2), "", !Branch2)
         Grid.Columns("Branch3").Value = IIf(IsNull(!Branch3), "", !Branch3)
         Grid.Columns("Branch4").Value = IIf(IsNull(!Branch4), "", !Branch4)
         Grid.Columns("Branch5").Value = IIf(IsNull(!Branch5), "", !Branch5)
         Grid.Columns("Branch6").Value = IIf(IsNull(!Branch6), "", !Branch6)
         Grid.Columns("Branch7").Value = IIf(IsNull(!Branch7), "", !Branch7)
         Grid.Columns("Branch8").Value = IIf(IsNull(!Branch8), "", !Branch8)
         Grid.Columns("Branch9").Value = IIf(IsNull(!Branch9), "", !Branch9)
       End If
        Grid.Columns("GroupID").Text = !GroupID
        Grid.Columns("packingname").Text = !PackingName
        Grid.Columns("multiplier").Value = !Multiplier
        Grid.Columns("Stock").Value = !Stock
        Grid.Columns("PurPrice").Value = !PurPrice
        Grid.Columns("ListPrice").Value = IIf(IsNull(!ListPrice), "", !ListPrice)
        Grid.Columns("WSPrice").Value = IIf(IsNull(!WSPrice), "", !WSPrice)
        Grid.Columns("RetailPrice").Value = !RetailPrice
        Grid.Columns("DiscPC").Value = IIf(IsNull(!DiscPC), "", !DiscPC)
        Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), "", !DiscPer)
        Grid.Columns("QtyLoose").Text = ""
        Grid.Columns("Group").Text = !GroupName
        Grid.Columns("SubGroup").Text = !SubGroupName
        Grid.Columns("Company").Text = !CompanyName
        Grid.Columns("Brand").Text = !BrandName
        If RdoSaleQty.Value = True Then
            Grid.Columns("QtyPack").Value = CN.Execute("SELECT dbo.FunGetPack(" & Val(!Productid) & ",Floor(" & !QtyLoose & "))").Fields(0).Value
            Grid.Columns("QtyLoose").Value = CN.Execute("SELECT dbo.FunGetLoose(" & Val(!Productid) & ",(" & !QtyLoose & "))").Fields(0).Value
            Grid.Update
            UpdateRs
         ElseIf RdoStockSale.Value = True Then
            Grid.Columns("QtyPack").Value = CN.Execute("SELECT dbo.FunGetPack(" & !Productid & ",Floor(" & !QtyLoose - !StockLoose & "))").Fields(0).Value
            Grid.Columns("QtyLoose").Value = CN.Execute("SELECT dbo.FunGetLoose(" & !Productid & ",(" & (!QtyLoose - !StockLoose) & "))").Fields(0).Value
            Grid.Update
            UpdateRs
         Else
            Grid.Update
         End If
        .MoveNext
      Loop
   End With
   vSuppressUpdateEvent = False
   Grid.Redraw = True
   Grid.MoveFirst
   Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnApply_Click()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.Rows - 1
      If Trim(TxtPurPrice.Text) <> "" Then
         Grid.Columns("PurPrice").Value = Val(TxtPurPrice.Text)
      End If
      If Trim(TxtQtyLoose.Text) <> "" Then
         Grid.Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
      End If
      If Trim(TxtDiscPer.Text) <> "" Then
         Grid.Columns("DiscPer").Text = TxtDiscPer.Text
         'Grid.Columns("DiscPC").Text = Val(Grid.Columns("RetailPrice").Text) - Val(TxtDiscPer.Text) / 100 * Val(Grid.Columns("RetailPer").Text)
         Grid.Columns("PurPrice").Text = Val(Grid.Columns("RetailPrice").Text) - Val(TxtDiscPer.Text) / 100 * Val(Grid.Columns("RetailPer").Text)
      End If
      Grid.MoveNext
   Next i
   UpdateRs
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnApplyStock_Click()
   On Error GoTo ErrorHandler
   Dim VStrSQL As String, vQtyLoose As Double
   
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.Rows - 1
      VStrSQL = "select isnull(dbo.FunStock(" & Val(Grid.Columns("ID").Value) & ",1,0,0,0,0,0,0,'" & GetServerDate + 1 & "',0),0)"
       With CN.Execute(VStrSQL)
         If .RecordCount > 0 Then
            vQtyLoose = .Fields(0).Value
         Else
            vQtyLoose = 0
         End If
         .Close
      End With
      Grid.Columns("QtyPack").Value = CN.Execute("SELECT dbo.FunGetPack(" & Val(Grid.Columns("ID").Value) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
      Grid.Columns("QtyLoose").Value = CN.Execute("SELECT dbo.FunGetLoose(" & Val(Grid.Columns("ID").Value) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
      Grid.MoveNext
   Next i
   UpdateRs
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
'  Call PopulateGrid
End Sub

Private Sub BtnClose_Click()
   Set RsTemp = New ADODB.Recordset
   RsTemp.Fields.Append "ProductID", adVarChar, 5
   RsTemp.Fields.Append "ItemCode", adVarChar, 9
   RsTemp.Fields.Append "ProductName", adVarChar, 100
   RsTemp.Fields.Append "Price", adDouble
   RsTemp.Fields.Append "QtyPack", adDouble
   RsTemp.Fields.Append "QtyLoose", adDouble
   RsTemp.Fields.Append "PackingName", adVarChar, 100
   RsTemp.Fields.Append "Multiplier", adDouble
   RsTemp.Fields.Append "PartyID", adVarChar, 20
   RsTemp.Fields.Append "DiscPer", adDouble
   RsTemp.Fields.Append "DiscPC", adDouble
   RsTemp.Fields.Append "RetailPrice", adDouble
   RsTemp.Open
   Unload Me
End Sub

Private Sub BtnFilter_Click()
   On Error GoTo ErrorHandler
   vProductID = IIf(Val(TxtToProductID.Text) = 0, IIf(Val(TxtFromProductID.Text) = 0, "", " and P.ProductID = " & Val(TxtFromProductID.Text)), IIf(Val(TxtFromProductID.Text) = 0, "", " and P.ProductID Between " & Val(TxtFromProductID.Text) & " and " & Val(TxtToProductID.Text)))
   vWords = Split(TxtProductName.Text, " ")
   vProductName = ""
   For i = 0 To UBound(vWords)
       vProductName = vProductName & " and Productname like '%" & Replace(vWords(i), "'", "''") & "%'"
   Next
   PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnFromProduct_Click()
   If FunSelectFromProduct(ssButton, True) = True Then
      TxtToProductID.SetFocus
   Else
      TxtFromProductID.SetFocus
   End If
End Sub

Private Sub BtnSelect_Click()
   On Error GoTo ErrorHandler
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnToProduct_Click()
   If FunSelectToProduct(ssButton, True) = True Then
      TxtProductName.SetFocus
   Else
      TxtToProductID.SetFocus
   End If
End Sub

Private Sub BtnUpdateStock_Click()
On Error GoTo ErrorHandler
    Me.MousePointer = vbHourglass
    CN.Execute ("ProdUpdatCurrentStockStore")
    CN.Execute ("ProdUpdatCurrentStock")
    Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnVender_Click()
   If FunSelectVender(ssButton, False) = True Then
      TxtFromProductID.SetFocus
   Else
      TxtVenderID.SetFocus
   End If
End Sub

Private Function FunSelectVender(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInDetail = ""
        SchAccounts.CmbFilter = "Vendors"
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectVender = False: Exit Function
        TxtVenderID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    VStrSQL = " Select c.* FROM ChartofAccounts c " & vbCrLf & _
              " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
              " where BarCode = '" & (TxtVenderID.Text) & "' or (c.AccountNo = " & Val(TxtVenderID.Text) & " and (c.AccountNo like '6%') and c.isDetailed = 1 and c.isLocked = 0)"
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtVenderID.Text = !AccountNo
          TxtVenderName.Text = !AccountName
          FunSelectVender = True
          .Close
          Exit Function
      Else
          FunSelectVender = False
          .Close
          TxtVenderID.Text = ""
          TxtVenderName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub ChkSaleDataStock_Click()
  DtpFrom.Visible = ChkSaleDataStock.Value = 1
   DtpTo.Visible = ChkSaleDataStock.Value = 1
   LblFrom.Visible = ChkSaleDataStock.Value = 1
   LblTo.Visible = ChkSaleDataStock.Value = 1
   FrmShowQty.Visible = ChkSaleDataStock.Value
   RdoNone.Value = Not ChkSaleDataStock.Value
End Sub

Private Sub TxtVenderID_Change()
   If TxtVenderID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtVenderID.Name Then Exit Sub
   If TxtVenderName.Text <> "" Then TxtVenderName.Text = ""
End Sub

Private Sub TxtVenderID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtVenderID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtVenderID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectVender(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectVender(ssButton, False)
   End If
   Cancel = vTemp
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
   SetWindowText Me.hWnd, "Product Range Grid"
   HelpLocation Me
 
   CmbSortBy.Clear
   CmbSortBy.AddItem "P.ProductID"
   CmbSortBy.AddItem "ProductName"
   CmbSortBy.AddItem "QtyLoose"
   
   If ChkSaleDataStock.Value = True Then
      FrmShowQty.Visible = True
   Else
      FrmShowQty.Visible = False
   End If
   
   If ObjRegistry.ShowMultiBranches = True Then
      vSQL = "SELECT 'Braches' AS Branches," & vbCrLf _
            + "[1] Branch1, [2] Branch2, [3] Branch3, [4] Branch4, [5] Branch5, [6] Branch6, [7] Branch7, [8] Branch8, [9] Branch9," & vbCrLf _
            + "S1.StoreName BranchName1, S2.StoreName BranchName2, S3.StoreName BranchName3, S4.StoreName BranchName4, S5.StoreName BranchName5, S6.StoreName BranchName6, S7.StoreName BranchName7, S8.StoreName BranchName8, S9.StoreName BranchName9" & vbCrLf _
            + "FROM (" & vbCrLf _
            + "SELECT StoreID FROM Stores" & vbCrLf _
            + ") AS SourceTable" & vbCrLf _
            + "PIVOT(" & vbCrLf _
            + "AVG(StoreID) FOR [StoreID] IN([1], [2], [3], [4], [5], [6], [7], [8], [9])" & vbCrLf _
            + ") AS PivotStore" & vbCrLf _
            + "Left Outer Join Stores S1 on S1.StoreID = PivotStore.[1]" & vbCrLf _
            + "Left Outer Join Stores S2 on S2.StoreID = PivotStore.[2]" & vbCrLf _
            + "Left Outer Join Stores S3 on S3.StoreID = PivotStore.[3]" & vbCrLf _
            + "Left Outer Join Stores S4 on S4.StoreID = PivotStore.[4]" & vbCrLf _
            + "Left Outer Join Stores S5 on S5.StoreID = PivotStore.[5]" & vbCrLf _
            + "Left Outer Join Stores S6 on S6.StoreID = PivotStore.[6]" & vbCrLf _
            + "Left Outer Join Stores S7 on S7.StoreID = PivotStore.[7]" & vbCrLf _
            + "Left Outer Join Stores S8 on S8.StoreID = PivotStore.[8]" & vbCrLf _
            + "Left Outer Join Stores S9 on S9.StoreID = PivotStore.[9]"
            
      With CN.Execute(vSQL)
         If Not .EOF Then
            If Not (IsNull(!Branch1)) Then
               Grid.Columns("Branch1").Visible = True
               Grid.Columns("Branch1").Width = 80
               Grid.Columns("Branch1").Caption = IIf(IsNull(!BranchName1), "", !BranchName1)
            End If
            If Not (IsNull(!Branch2)) Then
               Grid.Columns("Branch2").Visible = True
               Grid.Columns("Branch2").Width = 70
               Grid.Columns("Branch2").Caption = IIf(IsNull(!BranchName2), "", !BranchName2)
            End If
            If Not (IsNull(!Branch3)) Then
               Grid.Columns("Branch3").Visible = True
               Grid.Columns("Branch3").Width = 80
               Grid.Columns("Branch3").Caption = IIf(IsNull(!BranchName3), "", !BranchName3)
            End If
            If Not (IsNull(!Branch4)) Then
               Grid.Columns("Branch4").Visible = True
               Grid.Columns("Branch4").Width = 80
               Grid.Columns("Branch4").Caption = IIf(IsNull(!BranchName4), "", !BranchName4)
            End If
            If Not (IsNull(!Branch5)) Then
               Grid.Columns("Branch5").Visible = True
               Grid.Columns("Branch5").Width = 80
               Grid.Columns("Branch5").Caption = IIf(IsNull(!BranchName5), "", !BranchName5)
            End If
            If Not (IsNull(!Branch6)) Then
               Grid.Columns("Branch6").Visible = True
               Grid.Columns("Branch6").Width = 80
               Grid.Columns("Branch6").Caption = IIf(IsNull(!BranchName6), "", !BranchName6)
            End If
            If Not (IsNull(!Branch7)) Then
               Grid.Columns("Branch7").Visible = True
               Grid.Columns("Branch7").Width = 80
               Grid.Columns("Branch7").Caption = IIf(IsNull(!BranchName8), "", !BranchName7)
            End If
            If Not (IsNull(!Branch8)) Then
               Grid.Columns("Branch8").Visible = True
               Grid.Columns("Branch8").Width = 80
               Grid.Columns("Branch8").Caption = IIf(IsNull(!BranchName8), "", !BranchName8)
            End If
            If Not (IsNull(!Branch9)) Then
               Grid.Columns("Branch9").Visible = True
               Grid.Columns("Branch9").Width = 80
               Grid.Columns("Branch9").Caption = IIf(IsNull(!BranchName9), "", !BranchName9)
            End If
            
         End If
      End With
   End If

'   LblCompany.Visible = ObjRegistry.ShowAllPrices
'   BtnCompany.Visible = ObjRegistry.ShowAllPrices
'   TxtCompanyID.Visible = ObjRegistry.ShowAllPrices
'   TxtCompanyName.Visible = ObjRegistry.ShowAllPrices
   
   Grid.Columns("ListPrice").Visible = ObjRegistry.ShowAllPrices
   Grid.Columns("WSPrice").Visible = ObjRegistry.ShowAllPrices
   
   
'   CmbBrand.Clear
'   With cn.Execute("Select * FROM Brands Order By BrandName")
'      CmbBrand.AddItem "All Brands"
'      CmbBrand.ItemData(CmbBrand.NewIndex) = 0
'      Do Until .EOF
'         CmbBrand.AddItem !BrandName
'         CmbBrand.ItemData(CmbBrand.NewIndex) = !BrandID
'         .MoveNext
'      Loop
'   End With
'
'   CmbCompany.Clear
'   With cn.Execute("Select * FROM Companies Order By CompanyName")
'      CmbCompany.AddItem "All Companies"
'      CmbCompany.ItemData(CmbCompany.NewIndex) = 0
'      Do Until .EOF
'         CmbCompany.AddItem !CompanyName
'         CmbCompany.ItemData(CmbCompany.NewIndex) = !companyid
'         .MoveNext
'      Loop
'   End With
'
'   CmbGroup.Clear
'   With cn.Execute("Select * FROM Groups Order By GroupName")
'      CmbGroup.AddItem "All Groups"
'      CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left("000", 1)) & Asc(Mid("000", 2, 1)) & Asc(Mid("000", 3, 1))
'      Do Until .EOF
'         CmbGroup.AddItem !GroupName
'         CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left(!Groupid, 1)) & Asc(Mid(!Groupid, 2, 1)) & Asc(Mid(!Groupid, 3, 1))
'         .MoveNext
'      Loop
'   End With
'
'   CmbSubGroup.Clear
'   With cn.Execute("Select * FROM SubGroups Order By SubGroupName")
'      CmbSubGroup.AddItem "All SubGroups"
'      CmbSubGroup.ItemData(CmbSubGroup.NewIndex) = 0
'      Do Until .EOF
'         CmbSubGroup.AddItem !SubGroupName
'         CmbSubGroup.ItemData(CmbSubGroup.NewIndex) = !SubGroupID
'         .MoveNext
'      Loop
'   End With
'
'   CmbBrand.ListIndex = 0
'   CmbCompany.ListIndex = 0
'   CmbGroup.ListIndex = 0
'   CmbSubGroup.ListIndex = 0
   CmbSortBy.ListIndex = 1
   'PopulateGrid
   TxtVenderID.Text = ParaInPartyID
   If TxtVenderID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectVender(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectVender(ssButton, False)
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name <> Grid.Name Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnSelect.Enabled Then BtnSelect_Click
            KeyCode = 0
         Case vbKeyH
               FraHelp.ZOrder 0
               FraHelp.Visible = True
               KeyCode = 0
         Case vbKeyW
            If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then TxtVenderID.SetFocus
         Case TxtVenderID.Name: If FunSelectVender(ssFunctionKey, True) = True Then TxtFromProductID.SetFocus
         Case TxtFromProductID.Name: If FunSelectFromProduct(ssFunctionKey, True) = True Then TxtToProductID.SetFocus
         Case TxtToProductID.Name: If FunSelectToProduct(ssFunctionKey, True) = True Then TxtProductName.SetFocus
      End Select
   ElseIf ActiveControl.Name = Grid.Name And ObjRegistry.ShowAllPrices Then
      If ParaInBoth = False Then
         If Val(Grid.Columns("Multiplier").Text) = 0 Then
            If Grid.Col = 3 Then KeyCode = 0
         ElseIf Val(Grid.Columns("Multiplier").Text) <> 0 Then
            If Grid.Col = 4 Then KeyCode = 0
         End If
      End If
   End If
End Sub

Private Sub Grid_LostFocus()
   On Error GoTo ErrorHandler
   If vSuppressUpdateEvent Then Exit Sub
   If Grid.Visible = False Then Exit Sub
   UpdateRs
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  'If Grid.Columns(ColIndex).Text = "" Then Grid.Columns(ColIndex).Text = "0"
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   On Error GoTo ErrorHandler
   If vSuppressUpdateEvent Then Exit Sub
   UpdateRs
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_GotFocus()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> Grid.Name Then Exit Sub
   Grid.Row = 0
   Grid.Col = 0
   SendKeys "{Right}"
   Exit Sub
ErrorHandler:
   If err.Number = 70 Then Resume Next
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event vbKeyRight, 1, 1, 1
      KeyCode = 0
'   ElseIf Val(Grid.Columns("Multiplier").Text) = 0 Then
'      If Grid.Col = 3 Then KeyCode = 0
'   ElseIf Val(Grid.Columns("Multiplier").Text) <> 0 Then
'      If Grid.Col = 4 Then KeyCode = 0
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub UpdateRs()
   RsTemp.Filter = "ProductID = " & Val(Grid.Columns("ID").Text)
   If Grid.Col = 4 Or Grid.Col = 5 Then
'      Grid.Columns("PurPrice").Text = Val(Grid.Columns("RetailPrice").Text) - Val(Grid.Columns("RetailPrice").Text) / 100 * Val(Grid.Columns("RetailPer").Text)
   End If
   If RsTemp.RecordCount = 0 And ((Val(Grid.Columns("QtyLoose").Value) > 0 Or Val(Grid.Columns("QtyPack").Value) > 0) Or ParaInBoth = True) Then
      RsTemp.AddNew
      RsTemp!Productid = Val(Grid.Columns("ID").Text)
      RsTemp!ProductName = Grid.Columns("Name").Text
      RsTemp!GroupID = Grid.Columns("GroupID").Text
      RsTemp!PackingName = Grid.Columns("PackingName").Text
      RsTemp!QtyPack = Val(Grid.Columns("QtyPack").Text)
      RsTemp!Multiplier = Grid.Columns("Multiplier").Text
      RsTemp!QtyLoose = Val(Grid.Columns("QtyLoose").Text)
      RsTemp!DiscPC = Val(Grid.Columns("DiscPC").Text)
      RsTemp!DiscPer = Val(Grid.Columns("DiscPer").Text)
      RsTemp!Price = Val(Grid.Columns("PurPrice").Text)
      RsTemp!RetailPrice = Val(Grid.Columns("RetailPrice").Text)
      RsTemp!PartyID = TxtVenderID.Text
'   ElseIf RsTemp.RecordCount = 1 And (Val(Grid.Columns("QtyLoose").Value) + Val(Grid.Columns("QtyPack").Value) = 0) And Val(Grid.Columns("PurPrice").Value) = 0 Then
    ElseIf RsTemp.RecordCount = 1 And ((Val(Grid.Columns("QtyLoose").Value) + Val(Grid.Columns("QtyPack").Value) = 0) And ParaInBoth = False) Then
      RsTemp.Delete
   ElseIf RsTemp.RecordCount = 1 Then
      RsTemp!QtyPack = Val(Grid.Columns("QtyPack").Text)
      RsTemp!QtyLoose = Val(Grid.Columns("QtyLoose").Value)
      RsTemp!Price = Val(Grid.Columns("PurPrice").Value)
      RsTemp!DiscPC = Val(Grid.Columns("DiscPC").Text)
      RsTemp!DiscPer = Val(Grid.Columns("DiscPer").Text)
      RsTemp.Update
  End If
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtFromProductID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If Trim(TxtFromProductID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectFromProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectFromProduct(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectFromProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectFromProduct = False: Exit Function
      TxtFromProductID.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtFromProductID.Text) = "" Then Exit Function
    If TxtFromProductID.Text = "" Then FunSelectFromProduct = False: Exit Function
    
   VStrSQL = " SELECT p.ProductID, Code, Qty, ProductName" & vbCrLf _
         + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
         + " where (p.productid = " & Val(TxtFromProductID.Text) & " or Code = '" & TxtFromProductID.Text & "')" & " and isLocked = 0 "

  With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtFromProductID.Text = !Productid
         FunSelectFromProduct = True
         .Close
         Exit Function
      Else
         FunSelectFromProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtFromProductID.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtToProductID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtToProductID.Text) = "" Then Exit Sub
   vTemp = Not FunSelectToProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectToProduct(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectToProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectToProduct = False: Exit Function
      TxtToProductID.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtToProductID.Text) = "" Then Exit Function
    If TxtToProductID.Text = "" Then FunSelectToProduct = False: Exit Function
    VStrSQL = " SELECT p.Productid, ProductName" & vbCrLf _
           + " from Products p" & vbCrLf _
           + " where p.productid = " & Val(TxtToProductID.Text)
  
   With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtToProductID.Text = !Productid
         FunSelectToProduct = True
         .Close
         Exit Function
      Else
         FunSelectToProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtToProductID.Text = ""
         Exit Function
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub PopulateDataToHistoryGrid()
   On Error GoTo ErrorHandler
    Dim PSQL As String
    PSQL = "select top 1 isnull(b.Price,PurPrice) as Price from Products p left outer join Purchasebody b on p.ProductID = b.ProductID" & vbCrLf & _
    " left outer join PurchaseHeader h  on h.PurID = b.PurID and H.PurchaseDate = b.PurchaseDate " & vbCrLf & _
    " where p.productid = " & Val(Grid.Columns("ID").Text) & " and isnull(h.PurchaseDate,'01-01-2010')  < Getdate() order by b.PurchaseDate Desc"
    
    vSQL = "select top 1 isnull(b.Price,WSPrice) as Price from Products p left outer join Salebody b on p.ProductID = b.ProductID" & vbCrLf & _
    " left outer join SaleHeader h  on h.SID = b.SID " & vbCrLf & _
    " where p.productid = " & Val(Grid.Columns("ID").Text) & " and isnull(h.BillDate ,'01-01-2010') < Getdate() order by h.BillDate Desc"
    
    GridHistory.Redraw = False
    GridHistory.MoveFirst
    GridHistory.RemoveAll
    GridHistory.AllowAddNew = True
    GridHistory.AddNew
    With CN.Execute(PSQL)
        If .RecordCount > 0 Then GridHistory.Columns("PurPrice").Value = CN.Execute(PSQL).Fields(0).Value
        .Close
    End With
    With CN.Execute(vSQL)
        If .RecordCount > 0 Then GridHistory.Columns("WSPrice").Value = CN.Execute(vSQL).Fields(0).Value
        .Close
    End With
    GridHistory.Columns("Margin").Value = Val(GridHistory.Columns("WSPrice").Value) - Val(GridHistory.Columns("PurPrice").Value)
    If Val(GridHistory.Columns("WSPrice").Value) <> 0 Then
        GridHistory.Columns("MarginPer").Value = Round((Val(GridHistory.Columns("WSPrice").Value) - Val(GridHistory.Columns("PurPrice").Value)) / GridHistory.Columns("WSPrice").Value * 100, 2)
    End If
    GridHistory.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo ErrorHandler
   
   If ObjRegistry.ShowAllPrices Then
      PopulateDataToPriceGrid
      FrmProductPrices.Visible = True
      PopulateDataToHistoryGrid
      FrmHistory.Visible = True
      FrmHistory.ZOrder 0
      GridHistory.Visible = True
      GridHistory.ZOrder 0
   Else
      FrmProductPrices.Visible = False
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToPriceGrid()
    
      vSQL = "select desc1, Listprice, WsPrice, RetailPrice, PurPrice-isnull(purDiscPC*isnull(multiplier,1),0) as PurPrice " & vbCrLf & _
            " from products p" & vbCrLf & _
            " left outer join productpacking pp on p.productid = pp.productid and p.PurchasePackingID = pp.packingid" & vbCrLf & _
            " where p.productID = " & Val(Grid.Columns("ID").Text)
      
      
      With CN.Execute(vSQL)
         GridProductPrices.Redraw = False
         GridProductPrices.MoveFirst
         GridProductPrices.RemoveAll
         GridProductPrices.AllowAddNew = True
         While Not .EOF
            GridProductPrices.AddNew
            GridProductPrices.Columns("Description").Text = IIf(IsNull(!Desc1), "", !Desc1)
            GridProductPrices.Columns("Pur").Value = !PurPrice
            GridProductPrices.Columns("List").Value = IIf(IsNull(!ListPrice), "0", !ListPrice)
            GridProductPrices.Columns("WS").Value = !WSPrice
            GridProductPrices.Columns("Retail").Value = !RetailPrice
            .MoveNext
         Wend
         .Close
         GridProductPrices.MoveFirst
         GridProductPrices.Redraw = True
      End With
End Sub

Private Function FunSelectCompany(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCompany.Show vbModal, Me
        If SchCompany.ParaOutCompanyID = "" Then FunSelectCompany = False: Exit Function
        TxtCompanyID.Text = SchCompany.ParaOutCompanyID
    End If
    '---------------------------
    VStrSQL = " Select * FROM Companies where CompanyID=" & Val(TxtCompanyID.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtCompanyName.Text = !CompanyName
          FunSelectCompany = True
          .Close
          Exit Function
      Else
          FunSelectCompany = False
          .Close
          TxtCompanyID.Text = ""
          TxtCompanyName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtCompanyID_Change()
   If TxtCompanyID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCompanyID.Name Then Exit Sub
   If TxtCompanyName.Text <> "" Then TxtCompanyName.Text = ""
End Sub

Private Sub TxtCompanyID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtCompanyID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCompanyID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectCompany(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectCompany(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnCompany_Click()
   If FunSelectCompany(ssButton, False) = True Then
      TxtBrandID.SetFocus
   Else
      TxtCompanyID.SetFocus
   End If
End Sub

Private Sub TxtGroupID_Change()
   If TxtGroupID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtGroupID.Name Then Exit Sub
   If TxtGroupName.Text <> "" Then TxtGroupName.Text = ""
End Sub

Private Sub TxtGroupID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtGroupID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtGroupID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectGroup(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectGroup(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectGroup(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchGroup.Show vbModal, Me
        If SchGroup.ParaOutGroupID = "" Then FunSelectGroup = False: Exit Function
        TxtGroupID.Text = SchGroup.ParaOutGroupID
    End If
    '---------------------------
    If Trim(TxtGroupID.Text) = "" Then Exit Function
    If Len(TxtGroupID.Text) <= 3 Then
      TxtGroupID.Text = Right("000" + CStr(Val(TxtGroupID.Text)), 3)
    End If
    If TxtGroupID.Text = "" Then FunSelectGroup = False: Exit Function
    VStrSQL = " Select * FROM Groups where GroupID = '" & TxtGroupID.Text & "'"
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtGroupName.Text = !GroupName
          FunSelectGroup = True
          .Close
          Exit Function
      Else
          FunSelectGroup = False
          .Close
          TxtGroupID.Text = ""
          TxtGroupName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnGroup_Click()
   If FunSelectGroup(ssButton, False) = True Then
      TxtSubGroupID.SetFocus
   Else
      TxtGroupID.SetFocus
   End If
End Sub

Private Sub TxtSubGroupID_Change()
   If TxtSubGroupID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSubGroupID.Name Then Exit Sub
   If TxtSubGroupName.Text <> "" Then TxtSubGroupName.Text = ""
End Sub

Private Sub TxtSubGroupID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtSubGroupID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSubGroupID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSubGroup(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSubGroup(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSubGroup(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSubGroup.Show vbModal, Me
        If SchSubGroup.ParaOutSubGroupID = "" Then FunSelectSubGroup = False: Exit Function
        TxtSubGroupID.Text = SchSubGroup.ParaOutSubGroupID
    End If
    '---------------------------
    VStrSQL = " Select * FROM SubGroups where SubGroupID = " & Val(TxtSubGroupID.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtSubGroupName.Text = !SubGroupName
          FunSelectSubGroup = True
          .Close
          Exit Function
      Else
          FunSelectSubGroup = False
          .Close
          TxtSubGroupID.Text = ""
          TxtSubGroupName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnSubGroup_Click()
   If FunSelectSubGroup(ssButton, False) = True Then
      TxtBrandID.SetFocus
   Else
      TxtSubGroupID.SetFocus
   End If
End Sub
Private Sub BtnBrand_Click()
   If FunSelectBrand(ssButton, False) = True Then
      TxtGroupID.SetFocus
   Else
      TxtBrandID.SetFocus
   End If
End Sub

Private Function FunSelectBrand(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchBrand.Show vbModal, Me
        If SchBrand.ParaOutBrandID = "" Then FunSelectBrand = False: Exit Function
        TxtBrandID.Text = SchBrand.ParaOutBrandID
    End If
    '---------------------------
    VStrSQL = " Select * FROM Brands where BrandID=" & Val(TxtBrandID.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtBrandName.Text = !BrandName
          FunSelectBrand = True
          .Close
          Exit Function
      Else
          FunSelectBrand = False
          .Close
          TxtBrandID.Text = ""
          TxtBrandName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtBrandID_Change()
   If TxtBrandID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtBrandID.Name Then Exit Sub
   If TxtBrandName.Text <> "" Then TxtBrandName.Text = ""
End Sub

Private Sub TxtBrandID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtBrandID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtBrandID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectBrand(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectBrand(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

