VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmOpeningProductGrid 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmOpeningGrid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbBrand 
      Height          =   315
      Left            =   11295
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3030
      Width           =   1665
   End
   Begin VB.ComboBox CmbSubGroup 
      Height          =   315
      Left            =   9630
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3030
      Width           =   1665
   End
   Begin VB.TextBox TxtToProductID 
      Height          =   345
      Left            =   3810
      TabIndex        =   2
      Top             =   2370
      Width           =   1380
   End
   Begin VB.ComboBox CmbSortBy 
      Height          =   315
      Left            =   11310
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2355
      Width           =   1170
   End
   Begin VB.TextBox TxtFromProductID 
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Top             =   2370
      Width           =   1380
   End
   Begin VB.TextBox TxtProductName 
      Height          =   345
      Left            =   2040
      TabIndex        =   3
      Top             =   3000
      Width           =   2700
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   6045
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3030
      Width           =   1890
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   7950
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3030
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
      Top             =   1080
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
         Text            =   "FrmOpeningGrid.frx":0ECA
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
      Top             =   9120
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
      MICON           =   "FrmOpeningGrid.frx":0F1C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7170
      TabIndex        =   10
      Top             =   9120
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
      MICON           =   "FrmOpeningGrid.frx":0F38
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8460
      TabIndex        =   11
      Top             =   9120
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
      MICON           =   "FrmOpeningGrid.frx":0F54
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnFilter 
      Height          =   315
      Left            =   4815
      TabIndex        =   17
      Top             =   2985
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
      MICON           =   "FrmOpeningGrid.frx":0F70
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnApply 
      Height          =   315
      Left            =   10170
      TabIndex        =   18
      Top             =   2400
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
      MICON           =   "FrmOpeningGrid.frx":0F8C
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtQtyLoose 
      Height          =   315
      Left            =   9330
      TabIndex        =   19
      Top             =   2400
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
      Left            =   1890
      TabIndex        =   29
      Top             =   3480
      Width           =   11580
      ScrollBars      =   3
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   14
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
      stylesets(0).Picture=   "FrmOpeningGrid.frx":0FA8
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
      stylesets(1).Picture=   "FrmOpeningGrid.frx":0FC4
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
      Columns.Count   =   14
      Columns(0).Width=   1614
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ID"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   5768
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1323
      Columns(2).Caption=   "Stock"
      Columns(2).Name =   "Stock"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1085
      Columns(3).Caption=   "Qty (L)"
      Columns(3).Name =   "QtyLoose"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1429
      Columns(4).Caption=   "QtyPack"
      Columns(4).Name =   "QtyPack"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1746
      Columns(5).Caption=   "Retail Price"
      Columns(5).Name =   "RetailPrice"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1376
      Columns(6).Caption=   "Pur Price"
      Columns(6).Name =   "PurPrice"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   5
      Columns(6).FieldLen=   256
      Columns(7).Width=   1191
      Columns(7).Caption=   "Disc %"
      Columns(7).Name =   "RetailPer"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   2223
      Columns(8).Caption=   "Packing Name"
      Columns(8).Name =   "PackingName"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1376
      Columns(9).Caption=   "Multiplier"
      Columns(9).Name =   "multiplier"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   2461
      Columns(10).Caption=   "Company"
      Columns(10).Name=   "Company"
      Columns(10).CaptionAlignment=   2
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).NumberFormat=   "########.##"
      Columns(10).FieldLen=   256
      Columns(10).Locked=   -1  'True
      Columns(11).Width=   2461
      Columns(11).Caption=   "Group"
      Columns(11).Name=   "Group"
      Columns(11).CaptionAlignment=   2
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(11).Locked=   -1  'True
      Columns(12).Width=   2461
      Columns(12).Caption=   "Sub Group"
      Columns(12).Name=   "SubGroup"
      Columns(12).CaptionAlignment=   2
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(12).Locked=   -1  'True
      Columns(13).Width=   2461
      Columns(13).Caption=   "Brand"
      Columns(13).Name=   "Brand"
      Columns(13).CaptionAlignment=   2
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20426
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
      Left            =   3450
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2370
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
      MICON           =   "FrmOpeningGrid.frx":0FE0
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnToProduct 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5220
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2355
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
      MICON           =   "FrmOpeningGrid.frx":0FFC
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPurPrice 
      Height          =   315
      Left            =   8415
      TabIndex        =   32
      Top             =   2400
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
      TabIndex        =   34
      Top             =   2400
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
      Left            =   8445
      TabIndex        =   36
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1785
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
      MICON           =   "FrmOpeningGrid.frx":1018
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtVenderName 
      Height          =   315
      Left            =   8805
      TabIndex        =   37
      Tag             =   "nc"
      Top             =   1785
      Width           =   3585
      _ExtentX        =   6324
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
      Left            =   7425
      TabIndex        =   0
      Top             =   1785
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
      TabIndex        =   40
      Top             =   2355
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
      MICON           =   "FrmOpeningGrid.frx":1034
      BC              =   12632256
      FC              =   0
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
      Left            =   7425
      TabIndex        =   39
      Top             =   1590
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
      Left            =   8820
      TabIndex        =   38
      Top             =   1590
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
      TabIndex        =   35
      Top             =   2175
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
      TabIndex        =   33
      Top             =   2175
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
      Left            =   11295
      TabIndex        =   28
      Top             =   2805
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
      Left            =   9630
      TabIndex        =   27
      Top             =   2805
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
      Left            =   3810
      TabIndex        =   26
      Top             =   2130
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
      TabIndex        =   25
      Top             =   2175
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
      Left            =   11310
      TabIndex        =   24
      Top             =   2130
      Width           =   660
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
      Left            =   2040
      TabIndex        =   23
      Top             =   2130
      Width           =   1395
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
      Left            =   2040
      TabIndex        =   22
      Top             =   2760
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
      Left            =   6045
      TabIndex        =   21
      Top             =   2805
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
      Left            =   7950
      TabIndex        =   20
      Top             =   2805
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
      Caption         =   "Opening Product Grid"
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
      Width           =   2835
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "FrmOpeningProductGrid"
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
Dim vSQL As String
Public ParaInPartyID As String

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
   Me.MousePointer = vbHourglass
   
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
   RsTemp.Fields.Append "RetailPer", adDouble
   RsTemp.Fields.Append "RetailPrice", adDouble
   
   
   RsTemp.Open

   
   Me.MousePointer = vbHourglass
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
   
   If TxtVenderID.Text = "" Then
     vSQL = " SELECT P.ProductID, p.itemcode, ProductName, PurPrice, GroupName, isnull(CompanyName,'') as CompanyName, " & vbCrLf _
      + " isnull(SubGroupName,'') as SubGroupName, isnull(BrandName,'') as BrandName, RetailPrice," & vbCrLf _
      + " isnull(packingname,'') packingname, isnull(pp.multiplier,'') multiplier, CS.QtyLoose Stock" & vbCrLf _
      + " FROM Products p Left Outer join Groups g on g.GroupID = p.GroupID" & vbCrLf _
      + " Left outer join CurrentStock CS on cs.productid = p.productid" & vbCrLf _
      + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID" & vbCrLf _
      + " left outer join Companies c on c.CompanyID = p.CompanyID" & vbCrLf _
      + " Left outer join packings pk on pk.packingid = p.purchasepackingid" & vbCrLf _
      + " Left outer join productpacking pp on pp.productid = p.productid" & vbCrLf _
      + " left outer join Brands b on b.BrandID = p.BrandID" & vbCrLf _
      + " where p.isLocked = 0 " & vProductID & vProductName & vbCrLf _
      + IIf(CmbCompany.ListIndex = 0, "", " and p.CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex)) & vbCrLf _
      + IIf(CmbGroup.ListIndex = 0, "", " and p.GroupID ='" & GetGroupID(CmbGroup) & "'") & vbCrLf _
      + IIf(CmbSubGroup.ListIndex = 0, "", " and p.SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex)) & vbCrLf _
      + IIf(CmbBrand.ListIndex = 0, "", " and p.BrandID =" & CmbBrand.ItemData(CmbBrand.ListIndex)) & vbCrLf _
      + " Order by " + CmbSortBy.Text
   Else
'      vSQL = " SELECT P.ProductID, P.ItemCode, ProductName, PurPrice, GroupName, isnull(CompanyName,'') as CompanyName, " & vbCrLf _
      + " isnull(SubGroupName,'') as SubGroupName, isnull(BrandName,'') as BrandName, RetailPrice," & vbCrLf _
      + " isnull(packingname,'') packingname, isnull(pp.multiplier,'') multiplier, CS.QtyLoose Stock" & vbCrLf _
      + " FROM Products p Left Outer join Groups g on g.GroupID = p.GroupID" & vbCrLf _
      + " Left outer join CurrentStock CS on cs.productid = p.productid" & vbCrLf _
      + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID" & vbCrLf _
      + " left outer join Companies c on c.CompanyID = p.CompanyID" & vbCrLf _
      + " Left outer join packings pk on pk.packingid = p.purchasepackingid" & vbCrLf _
      + " Left outer join productpacking pp on pp.productid = p.productid" & vbCrLf _
      + " left outer join Brands b on b.BrandID = p.BrandID" & vbCrLf _
      + " Inner join (select distinct productid  from Purchaseheader h left outer join purchasebody b on  h.PurID = b.PurID and h.PurchaseDate = b.PurchaseDate where vendorid = '" & TxtVenderID.Text & "') pur on pur.productid = p.productid" & vbCrLf _
      + " where p.isLocked = 0 " & vProductID & vProductName & vbCrLf _
      + IIf(CmbCompany.ListIndex = 0, "", " and p.CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex)) & vbCrLf _
      + IIf(CmbGroup.ListIndex = 0, "", " and p.GroupID ='" & GetGroupID(CmbGroup) & "'") & vbCrLf _
      + IIf(CmbSubGroup.ListIndex = 0, "", " and p.SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex)) & vbCrLf _
      + IIf(CmbBrand.ListIndex = 0, "", " and p.BrandID =" & CmbBrand.ItemData(CmbBrand.ListIndex)) & vbCrLf _
      + " Order by " + CmbSortBy.Text
       vSQL = " SELECT P.ProductID, P.ItemCode, ProductName, PurPrice, GroupName, isnull(CompanyName,'') as CompanyName, " & vbCrLf _
      + " isnull(SubGroupName,'') as SubGroupName, isnull(BrandName,'') as BrandName, RetailPrice," & vbCrLf _
      + " isnull(packingname,'') packingname, isnull(pp.multiplier,'') multiplier, CS.QtyLoose Stock" & vbCrLf _
      + " FROM Products p Left Outer join Groups g on g.GroupID = p.GroupID" & vbCrLf _
      + " Left outer join CurrentStock CS on cs.productid = p.productid" & vbCrLf _
      + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID" & vbCrLf _
      + " left outer join Companies c on c.CompanyID = p.CompanyID" & vbCrLf _
      + " Left outer join packings pk on pk.packingid = p.purchasepackingid" & vbCrLf _
      + " Left outer join productpacking pp on pp.productid = p.productid" & vbCrLf _
      + " left outer join Brands b on b.BrandID = p.BrandID" & vbCrLf _
      + " Inner join (select distinct b.productid  from Purchaseheader h left outer join purchasebody b on  h.PurID = b.PurID and h.PurchaseDate = b.PurchaseDate " & vbCrLf _
      + " inner join  " & vbCrLf _
      + " (select ProductID, max(SerialNo) SerialNO from purchasebody group by ProductID) d  " & vbCrLf _
      + " on b.SerialNo = d.SerialNo " & vbCrLf _
      + " where vendorid = '" & TxtVenderID.Text & "') pur on pur.productid = p.productid" & vbCrLf _
      + " where p.isLocked = 0 " & vProductID & vProductName & vbCrLf _
      + IIf(CmbCompany.ListIndex = 0, "", " and p.CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex)) & vbCrLf _
      + IIf(CmbGroup.ListIndex = 0, "", " and p.GroupID ='" & GetGroupID(CmbGroup) & "'") & vbCrLf _
      + IIf(CmbSubGroup.ListIndex = 0, "", " and p.SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex)) & vbCrLf _
      + IIf(CmbBrand.ListIndex = 0, "", " and p.BrandID =" & CmbBrand.ItemData(CmbBrand.ListIndex)) & vbCrLf _
      + " Order by " + CmbSortBy.Text
   End If
  
   With cn.Execute(vSQL)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("packingname").Text = !Packingname
        Grid.Columns("multiplier").Value = !Multiplier
        Grid.Columns("Stock").Value = !Stock
        Grid.Columns("PurPrice").Value = !PurPrice
        Grid.Columns("RetailPrice").Value = !RetailPrice
        Grid.Columns("QtyLoose").Text = ""
        Grid.Columns("Group").Text = !GroupName
        Grid.Columns("SubGroup").Text = !SubGroupName
        Grid.Columns("Company").Text = !CompanyName
        Grid.Columns("Brand").Text = !BrandName
        Grid.Update
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
      If Trim(TxtQtyLoose.Text) <> "" Then
         Grid.Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
      End If
      If Trim(TxtDiscPer.Text) <> "" Then
         Grid.Columns("RetailPer").Text = TxtDiscPer.Text
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
   RsTemp.Fields.Append "RetailPer", adDouble
   RsTemp.Fields.Append "RetailPrice", adDouble
   RsTemp.Open
   Unload Me
End Sub

Private Sub BtnFilter_Click()
   On Error GoTo ErrorHandler
   vProductID = IIf(Val(TxtToProductID.Text) = 0, IIf(Val(TxtFromProductID.Text) = 0, "", " and P.ProductID = '" & TxtFromProductID.Text & "'"), IIf(Val(TxtFromProductID.Text) = 0, "", " and P.ProductID Between '" & TxtFromProductID.Text & "' and '" & TxtToProductID.Text & "'"))
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
    cn.Execute ("ProdUpdatCurrentStock")
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
    Dim vStrSQL As String
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
    vStrSQL = " Select c.* FROM ChartofAccounts c " & vbCrLf & _
              " Left Outer join Parties p on c.AccountNo = p.PartyID " & vbCrLf & _
              " where BarCode = '" & (TxtVenderID.Text) & "' or (c.AccountNo = '" & (TxtVenderID.Text) & "' and (c.AccountNo like '6%') and c.isDetailed = 1 and c.isLocked = 0)"
    With cn.Execute(vStrSQL)
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
   
   CmbBrand.Clear
   With cn.Execute("Select * FROM Brands Order By BrandName")
      CmbBrand.AddItem "All Brands"
      CmbBrand.ItemData(CmbBrand.NewIndex) = 0
      Do Until .EOF
         CmbBrand.AddItem !BrandName
         CmbBrand.ItemData(CmbBrand.NewIndex) = !BrandID
         .MoveNext
      Loop
   End With
   
   CmbCompany.Clear
   With cn.Execute("Select * FROM Companies Order By CompanyName")
      CmbCompany.AddItem "All Companies"
      CmbCompany.ItemData(CmbCompany.NewIndex) = 0
      Do Until .EOF
         CmbCompany.AddItem !CompanyName
         CmbCompany.ItemData(CmbCompany.NewIndex) = !companyid
         .MoveNext
      Loop
   End With
   
   CmbGroup.Clear
   With cn.Execute("Select * FROM Groups Order By GroupName")
      CmbGroup.AddItem "All Groups"
      CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left("000", 1)) & Asc(Mid("000", 2, 1)) & Asc(Mid("000", 3, 1))
      Do Until .EOF
         CmbGroup.AddItem !GroupName
         CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
         .MoveNext
      Loop
   End With
     
   CmbSubGroup.Clear
   With cn.Execute("Select * FROM SubGroups Order By SubGroupName")
      CmbSubGroup.AddItem "All SubGroups"
      CmbSubGroup.ItemData(CmbSubGroup.NewIndex) = 0
      Do Until .EOF
         CmbSubGroup.AddItem !SubGroupName
         CmbSubGroup.ItemData(CmbSubGroup.NewIndex) = !SubGroupID
         .MoveNext
      Loop
   End With
     
   CmbBrand.ListIndex = 0
   CmbCompany.ListIndex = 0
   CmbGroup.ListIndex = 0
   CmbSubGroup.ListIndex = 0
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
         Case TxtFromProductID.Name: If FunSelectFromProduct(ssFunctionKey, True) = True Then TxtToProductID.SetFocus
         Case TxtToProductID.Name: If FunSelectToProduct(ssFunctionKey, True) = True Then TxtProductName.SetFocus
      End Select
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
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event vbKeyRight, 1, 1, 1
      KeyCode = 0
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub UpdateRs()
   RsTemp.Filter = "ProductID='" & Grid.Columns("ID").Text & "'"
   If Grid.Col = 4 Or Grid.Col = 5 Then
'      Grid.Columns("PurPrice").Text = Val(Grid.Columns("RetailPrice").Text) - Val(Grid.Columns("RetailPrice").Text) / 100 * Val(Grid.Columns("RetailPer").Text)
   End If
   If RsTemp.RecordCount = 0 And (Val(Grid.Columns("QtyLoose").Value) > 0 Or Val(Grid.Columns("QtyPack").Value) > 0) Then
      RsTemp.AddNew
      RsTemp!Productid = Grid.Columns("ID").Text
      RsTemp!ProductName = Grid.Columns("Name").Text
      RsTemp!Packingname = Grid.Columns("PackingName").Text
      RsTemp!QtyPack = Val(Grid.Columns("QtyPack").Text)
      RsTemp!Multiplier = Grid.Columns("Multiplier").Text
      RsTemp!QtyLoose = Val(Grid.Columns("QtyLoose").Text)
'      RsTemp!RetailPer = Val(Grid.Columns("RetailPer").Text)
      RsTemp!Price = Val(Grid.Columns("PurPrice").Text)
      RsTemp!RetailPrice = Val(Grid.Columns("RetailPrice").Text)
      RsTemp!PartyID = TxtVenderID.Text

      
   ElseIf RsTemp.RecordCount = 1 And (Val(Grid.Columns("QtyLoose").Value) + Val(Grid.Columns("QtyPack").Value) = 0) And Val(Grid.Columns("PurPrice").Value) = 0 Then
      RsTemp.Delete
   ElseIf RsTemp.RecordCount = 1 Then
      RsTemp!QtyPack = Val(Grid.Columns("QtyPack").Text)
      RsTemp!QtyLoose = Val(Grid.Columns("QtyLoose").Value)
      RsTemp!Price = Val(Grid.Columns("PurPrice").Value)
'      RsTemp!RetailPer = Val(Grid.Columns("RetailPer").Text)
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
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectFromProduct = False: Exit Function
      TxtFromProductID.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtFromProductID.Text) = "" Then Exit Function
    If Len(TxtFromProductID.Text) <= 5 Then
      TxtFromProductID.Text = Right("00000" + CStr(Val(TxtFromProductID.Text)), 5)
    End If
    If TxtFromProductID.Text = "" Then FunSelectFromProduct = False: Exit Function
    
   vStrSQL = " SELECT p.ProductID, Code, Qty, ProductName" & vbCrLf _
         + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
         + " where (p.productid = '" & TxtFromProductID.Text & "' or Code = '" & TxtFromProductID.Text & "')" & " and isLocked = 0 "

  With cn.Execute(vStrSQL)
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
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectToProduct = False: Exit Function
      TxtToProductID.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtToProductID.Text) = "" Then Exit Function
    If Len(TxtToProductID.Text) <= 5 Then
      TxtToProductID.Text = Right("00000" + CStr(Val(TxtToProductID.Text)), 5)
    End If
    If TxtToProductID.Text = "" Then FunSelectToProduct = False: Exit Function
    vStrSQL = " SELECT p.Productid, ProductName" & vbCrLf _
           + " from Products p" & vbCrLf _
           + " where p.productid = '" & TxtToProductID.Text & "'"
  
   With cn.Execute(vStrSQL)
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

