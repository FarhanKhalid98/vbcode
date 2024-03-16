VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmChangePrice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmChangePrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmChangePrice.frx":0ECA
      Left            =   9315
      List            =   "FrmChangePrice.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Tag             =   "1"
      Top             =   10080
      Width           =   3585
   End
   Begin VB.CheckBox ChkShowLocikProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Include Locik Product"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   12780
      TabIndex        =   45
      Top             =   2250
      Width           =   2160
   End
   Begin VB.CheckBox ChkShowEmptyPCTCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Show Empty PCT Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10575
      TabIndex        =   44
      Top             =   2250
      Width           =   2205
   End
   Begin VB.CheckBox chkSearchAllProductName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Search All Product Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13005
      TabIndex        =   43
      Top             =   450
      Value           =   1  'Checked
      Width           =   2160
   End
   Begin VB.CheckBox Chk3rdScheduleItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "3rd Schedule Item"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   12420
      TabIndex        =   41
      Top             =   2025
      Width           =   1620
   End
   Begin VB.Frame FrmHistory 
      Height          =   1275
      Left            =   4905
      TabIndex        =   33
      Top             =   630
      Width           =   4230
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridHistory 
         Height          =   1050
         Left            =   45
         TabIndex        =   34
         Top             =   135
         Width           =   4110
         ScrollBars      =   2
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
         stylesets(0).Picture=   "FrmChangePrice.frx":0ECE
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
         stylesets(1).Picture=   "FrmChangePrice.frx":0EEA
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
         stylesets(2).Picture=   "FrmChangePrice.frx":0F06
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
         Columns.Count   =   5
         Columns(0).Width=   1402
         Columns(0).Caption=   "PurPrice"
         Columns(0).Name =   "PurPrice"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1508
         Columns(1).Caption=   "WS Price"
         Columns(1).Name =   "WSPrice"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1508
         Columns(2).Caption=   "R Price"
         Columns(2).Name =   "RetailPrice"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1138
         Columns(3).Caption=   "Margin"
         Columns(3).Name =   "Margin"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   1164
         Columns(4).Caption=   "Mrgn %"
         Columns(4).Name =   "MarginPer"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   7250
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
   Begin VB.ComboBox CmbSubGroup 
      Height          =   315
      Left            =   7620
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2205
      Width           =   1665
   End
   Begin VB.ComboBox cmbDepartment 
      Height          =   315
      Left            =   5220
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   1350
      Width           =   1890
   End
   Begin VB.ComboBox cmbSubDepartment 
      Height          =   315
      Left            =   7110
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   1350
      Width           =   1890
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   5955
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2205
      Width           =   1665
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   4050
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2205
      Width           =   1890
   End
   Begin VB.TextBox TxtProductName 
      Height          =   345
      Left            =   1080
      TabIndex        =   4
      Top             =   2175
      Width           =   1755
   End
   Begin VB.TextBox TxtProductID 
      Height          =   345
      Left            =   135
      TabIndex        =   3
      Top             =   2175
      Width           =   930
   End
   Begin VB.ComboBox CmbSortBy 
      Height          =   315
      Left            =   9330
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2205
      Width           =   1170
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5850
      Left            =   135
      TabIndex        =   8
      Top             =   2520
      Width           =   15105
      ScrollBars      =   3
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   25
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
      stylesets(0).Picture=   "FrmChangePrice.frx":0F22
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
      stylesets(1).Picture=   "FrmChangePrice.frx":0F3E
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedCol"
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
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
      Columns.Count   =   25
      Columns(0).Width=   1138
      Columns(0).Caption=   "P ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "ItemCode"
      Columns(1).Name =   "ItemCode"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   5715
      Columns(2).Caption=   "Product Name"
      Columns(2).Name =   "Name"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1640
      Columns(3).Caption=   "Packing"
      Columns(3).Name =   "Packing"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   900
      Columns(4).Caption=   "Mul"
      Columns(4).Name =   "Multiplier"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).NumberFormat=   "########.##"
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   1402
      Columns(5).Caption=   "Pur Price"
      Columns(5).Name =   "PurPrice"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   5
      Columns(5).NumberFormat=   "########.##"
      Columns(5).FieldLen=   256
      Columns(6).Width=   1773
      Columns(6).Caption=   "List Price"
      Columns(6).Name =   "ListPrice"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1402
      Columns(7).Caption=   "T Price"
      Columns(7).Name =   "WSPrice"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1402
      Columns(8).Caption=   "R Price"
      Columns(8).Name =   "RetailPrice"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1111
      Columns(9).Caption=   "Mrgn%"
      Columns(9).Name =   "Margin"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1270
      Columns(10).Caption=   "Disc/PC"
      Columns(10).Name=   "DiscPC"
      Columns(10).Alignment=   1
      Columns(10).CaptionAlignment=   2
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   1111
      Columns(11).Caption=   "Disc %"
      Columns(11).Name=   "DiscPer"
      Columns(11).Alignment=   1
      Columns(11).CaptionAlignment=   2
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   1270
      Columns(12).Caption=   "DiscVal"
      Columns(12).Name=   "DiscVal"
      Columns(12).Alignment=   1
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   1217
      Columns(13).Caption=   "STax%"
      Columns(13).Name=   "SaletaxPer"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   1402
      Columns(14).Caption=   "PCTCode"
      Columns(14).Name=   "PCTCode"
      Columns(14).CaptionAlignment=   2
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(15).Width=   820
      Columns(15).Caption=   "3rdSchedule"
      Columns(15).Name=   "3rdSchedule"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(15).Style=   2
      Columns(16).Width=   1217
      Columns(16).Caption=   "Min Lt."
      Columns(16).Name=   "MinStockLimit"
      Columns(16).Alignment=   1
      Columns(16).CaptionAlignment=   2
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   5
      Columns(16).FieldLen=   256
      Columns(17).Width=   1217
      Columns(17).Caption=   "Max Lt."
      Columns(17).Name=   "MaxStockLimit"
      Columns(17).Alignment=   1
      Columns(17).CaptionAlignment=   2
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   8
      Columns(17).FieldLen=   256
      Columns(18).Width=   820
      Columns(18).Caption=   "Lock"
      Columns(18).Name=   "Lock"
      Columns(18).CaptionAlignment=   2
      Columns(18).DataField=   "Column 18"
      Columns(18).DataType=   8
      Columns(18).FieldLen=   256
      Columns(18).Style=   2
      Columns(19).Width=   1164
      Columns(19).Caption=   "NoCost"
      Columns(19).Name=   "NoCost"
      Columns(19).DataField=   "Column 19"
      Columns(19).DataType=   8
      Columns(19).FieldLen=   256
      Columns(19).Style=   2
      Columns(20).Width=   820
      Columns(20).Caption=   "Raw"
      Columns(20).Name=   "Raw"
      Columns(20).DataField=   "Column 20"
      Columns(20).DataType=   8
      Columns(20).FieldLen=   256
      Columns(20).Style=   2
      Columns(21).Width=   953
      Columns(21).Caption=   "Dead"
      Columns(21).Name=   "Dead"
      Columns(21).DataField=   "Column 21"
      Columns(21).DataType=   11
      Columns(21).FieldLen=   256
      Columns(21).Style=   2
      Columns(22).Width=   2566
      Columns(22).Caption=   "SaleMultiplier"
      Columns(22).Name=   "SaleMultiplier"
      Columns(22).DataField=   "Column 22"
      Columns(22).DataType=   8
      Columns(22).FieldLen=   256
      Columns(23).Width=   3200
      Columns(23).Caption=   "PurDiscPC"
      Columns(23).Name=   "PurDiscPC"
      Columns(23).DataField=   "Column 23"
      Columns(23).DataType=   8
      Columns(23).FieldLen=   256
      Columns(24).Width=   3200
      Columns(24).Caption=   "PurDiscPer"
      Columns(24).Name=   "PurDiscPer"
      Columns(24).DataField=   "Column 24"
      Columns(24).DataType=   8
      Columns(24).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   26644
      _ExtentY        =   10319
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
   Begin JeweledBut.JeweledButton BtnFilter 
      Height          =   315
      Left            =   2880
      TabIndex        =   5
      Top             =   2205
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
      MICON           =   "FrmChangePrice.frx":0F5A
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnApply 
      Height          =   315
      Left            =   14175
      TabIndex        =   7
      Top             =   855
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
      MICON           =   "FrmChangePrice.frx":0F76
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtDiscPer 
      Height          =   315
      Left            =   13365
      TabIndex        =   6
      Top             =   855
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
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6375
      TabIndex        =   16
      Top             =   8775
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
      MICON           =   "FrmChangePrice.frx":0F92
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7710
      TabIndex        =   17
      Top             =   8775
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
      MICON           =   "FrmChangePrice.frx":0FAE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   9045
      TabIndex        =   18
      Top             =   8775
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
      MICON           =   "FrmChangePrice.frx":0FCA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnMemberDisc 
      Height          =   420
      Left            =   5040
      TabIndex        =   19
      Top             =   8775
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Member Disc."
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
      MICON           =   "FrmChangePrice.frx":0FE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnNegativeMargin 
      Height          =   900
      Left            =   11340
      TabIndex        =   20
      Top             =   810
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   1588
      TX              =   "Show B/W From To Margins"
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
      MICON           =   "FrmChangePrice.frx":1002
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtFromMargin 
      Height          =   315
      Left            =   9465
      TabIndex        =   21
      Top             =   1320
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
   Begin SITextBox.Txt TxtToMargin 
      Height          =   315
      Left            =   10425
      TabIndex        =   23
      Top             =   1320
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   10
      Text            =   "-0.001"
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
   Begin SITextBox.Txt TxtItemCode 
      Height          =   315
      Left            =   3330
      TabIndex        =   25
      Top             =   1350
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
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
      Masked          =   1
      IntegralPoint   =   5
   End
   Begin JeweledBut.JeweledButton BtnApplySaleTax 
      Height          =   315
      Left            =   14175
      TabIndex        =   35
      Top             =   1215
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
      MICON           =   "FrmChangePrice.frx":101E
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtSaleTaxPer 
      Height          =   315
      Left            =   13365
      TabIndex        =   36
      Top             =   1215
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
   Begin JeweledBut.JeweledButton BtnApplyPCTCode 
      Height          =   315
      Left            =   14175
      TabIndex        =   38
      Top             =   1575
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
      MICON           =   "FrmChangePrice.frx":103A
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtPCTCode 
      Height          =   315
      Left            =   13365
      TabIndex        =   39
      Top             =   1575
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      Alignment       =   1
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
      IntegralPoint   =   8
   End
   Begin JeweledBut.JeweledButton BtnApply3rdScheduleItem 
      Height          =   315
      Left            =   14175
      TabIndex        =   42
      Top             =   1935
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
      MICON           =   "FrmChangePrice.frx":1056
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Height          =   420
      Left            =   3645
      TabIndex        =   46
      Top             =   8775
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
      MICON           =   "FrmChangePrice.frx":1072
      BC              =   14737632
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
      Left            =   8595
      TabIndex        =   48
      Top             =   10125
      Width           =   570
   End
   Begin VB.Label LblPCTCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PCT Code"
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
      Left            =   12420
      TabIndex        =   40
      Top             =   1665
      Width           =   870
   End
   Begin VB.Label LblSaleTaxPer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SaleTax %"
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
      Left            =   12390
      TabIndex        =   37
      Top             =   1305
      Width           =   900
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7620
      TabIndex        =   32
      Top             =   1980
      Width           =   1455
   End
   Begin VB.Label LblDepartment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Left            =   5220
      TabIndex        =   30
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label LblSubDepartment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Department"
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
      Left            =   7110
      TabIndex        =   29
      Top             =   1080
      Width           =   1380
   End
   Begin VB.Label LblItemCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
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
      Left            =   2385
      TabIndex        =   26
      Top             =   1395
      Width           =   870
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Margin"
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
      Left            =   10425
      TabIndex        =   24
      Top             =   1095
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Margin"
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
      Left            =   9225
      TabIndex        =   22
      Top             =   1095
      Width           =   1050
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
      Left            =   5955
      TabIndex        =   15
      Top             =   1980
      Width           =   1065
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
      Left            =   4050
      TabIndex        =   14
      Top             =   1980
      Width           =   1320
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
      Left            =   1080
      TabIndex        =   13
      Top             =   1935
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1935
      Width           =   930
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
      Left            =   9330
      TabIndex        =   11
      Top             =   1980
      Width           =   660
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc Per"
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
      Left            =   12555
      TabIndex        =   10
      Top             =   900
      Width           =   735
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Price"
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
      Height          =   360
      Index           =   0
      Left            =   2700
      TabIndex        =   9
      Top             =   270
      Width           =   1920
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "FrmChangePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Application1 As New CRAXDRT.Application
Dim RsReport As New ADODB.Recordset
Public Rs As New ADODB.Recordset
Public vSuppressUpdateEvent As Boolean
Public vParaSave As Boolean
Public vParaProductID As String
Public vParaSQL As String
Dim vMargin As String
Dim vShowBetween As String
Dim ssql, vSQL, vRandomID As String

Private Sub BtnApply_Click()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.Rows - 1
       Grid.Columns("DiscPer").Value = Val(TxtDiscPer.Text)
       If ObjRegistry.ShowWholeSaleMargin = True Then
         Grid.Columns("DiscPc").Value = Round((Val(Grid.Columns("WSPrice").Value) * Val(TxtDiscPer.Text) / 100), 2)
         Grid.Columns("DiscVal").Value = Round(Val(Grid.Columns("WSPrice").Value) * Val(TxtDiscPer.Text) / 100, 2)
       Else
         Grid.Columns("DiscPc").Value = Round((Val(Grid.Columns("RetailPrice").Value) * Val(TxtDiscPer.Text) / 100), 2)
         Grid.Columns("DiscVal").Value = Round(Val(Grid.Columns("RetailPrice").Value) * Val(TxtDiscPer.Text) / 100, 2)
       End If
       Grid.MoveNext
   Next i
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnApply3rdScheduleItem_Click()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.Rows - 1
       Grid.Columns("3rdSchedule").Value = Chk3rdScheduleItem.Value
       Grid.MoveNext
   Next i
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnApplySaleTax_Click()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.Rows - 1
       Grid.Columns("SaletaxPer").Value = Val(TxtSaleTaxPer.Text)
       Grid.MoveNext
   Next i
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnApplyPCTCode_Click()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.Rows - 1
       Grid.Columns("PCTCode").Value = Val(TxtPCTCode.Text)
       Grid.MoveNext
   Next i
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnFilter_Click()
   On Error GoTo ErrorHandler
   PopulateGrid
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnNegativeMargin_Click()
   On Error GoTo ErrorHandler
   'If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
Abc:
   Dim vStrSQL As String
   vStrSQL = " and round(case when RetailPrice = 0 then 0 else (isnull(RetailPrice,0) - ((PurPrice-isnull(purDiscPC*isnull(multiplier,1),0))/isnull(pp.multiplier,1)))*100/ isnull(RetailPrice,1) end,3) between " & IIf(Trim(TxtFromMargin.Text) = "", "-9999999", Val(TxtFromMargin.Text)) & " and " & IIf(Trim(TxtToMargin.Text) = "", "9999999", Val(TxtToMargin.Text))
   If ObjRegistry.ShowWholeSaleMargin = True Then
'      vMargin = "round(case when WSPrice = 0 then 0 else (isnull(WSPrice,0) - (PurPrice/isnull(sp.multiplier,1)))*100/ isnull(WSPrice,1) end,3) as margin"
'       vMargin = "round(case when WSPrice = 0 then 0 else (isnull(WSPrice,0) - PurPrice)*100/ isnull(WSPrice,1) end,3) as margin"
'       vShowBetween = " and round(case when WSPrice = 0 then 0 else (isnull(WSPrice,0) - PurPrice)*100/ isnull(WSPrice,1) end,3) between " & IIf(Trim(TxtFromMargin.Text) = "", "-9999999", Val(TxtFromMargin.Text)) & " and " & IIf(Trim(TxtToMargin.Text) = "", "9999999", Val(TxtToMargin.Text))
      vMargin = "round(case when WSPrice = 0 then 0 else (isnull(WSPrice,0)/isnull(sp.multiplier,1) - (PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0))/isnull(pp.multiplier,1))*100/ isnull(WSPrice,1) end,3) as margin"
      vShowBetween = " and round(case when WSPrice = 0 then 0 else (isnull(WSPrice,0)/isnull(sp.multiplier,1) - (PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0))/isnull(pp.multiplier,1))*100/ isnull(WSPrice,1) end,3) between " & IIf(Trim(TxtFromMargin.Text) = "", "-9999999", Val(TxtFromMargin.Text)) & " and " & IIf(Trim(TxtToMargin.Text) = "", "9999999", Val(TxtToMargin.Text))
  Else
      vMargin = "round(case when RetailPrice = 0 then 0 else (isnull(RetailPrice,0) - ((PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0))/isnull(pp.multiplier,1)))*100/ isnull(RetailPrice,1) end,3) as margin"
      vShowBetween = " and round(case when RetailPrice = 0 then 0 else (isnull(RetailPrice,0) - ((PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0))/isnull(pp.multiplier,1)))*100/ isnull(RetailPrice,1) end,3) between " & IIf(Trim(TxtFromMargin.Text) = "", "-9999999", Val(TxtFromMargin.Text)) & " and " & IIf(Trim(TxtToMargin.Text) = "", "9999999", Val(TxtToMargin.Text))
  End If
  If Rs.State = adStateOpen Then
    Rs.CancelBatch
    Rs.Close
  End If
  Me.MousePointer = vbHourglass
'  ssql = "Select distinct p.ProductID, ProductName, PurPrice, PurDiscPC, WSPrice, RetailPrice, DiscPC, DiscPer, SaleTaxPer, PCTCOde, is3rdScheduleItem, MinStockLimit, MaxStockLimit, isLocked, IsNoCostProduct, IsRawProduct FROM Products p left outer join ProductBarCodes b on p.productid = b.productid where 1=1 and PurPrice > RetailPrice " & IIf(TxtProductID.Text = "", "", " and p.ProductID = " & Val(TxtProductID.Text) & " or Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'")
   ssql = " SELECT distinct p.ProductID, ProductName, PurPrice, WSPrice, RetailPrice, ListPrice, PurDiscPer, PurDiscPC, DiscPC, DiscPer, SaleTaxPer, PCTCOde, is3rdScheduleItem, MinStockLimit, MaxStockLimit, isLocked, IsNoCostProduct, IsRawProduct, isDeadProduct, P.IsSync, modified_on from Products p" & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join ProductPacking SP on SP.packingid = P.SalePackingID and SP.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 " & vShowBetween & " and PurPrice > RetailPrice " & IIf(TxtProductID.Text = "", "", " and p.ProductID = " & Val(TxtProductID.Text) & " or Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'")
 
  Rs.Open ssql, CN, adOpenStatic, adLockBatchOptimistic
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
  
   ssql = " SELECT p.*, isnull(PackingName,'') as PackingName, isnull(pp.Multiplier,0) as Multiplier, isnull(SP.Multiplier,0) as SaleMultiplier, PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0) as PurchasePrice, isnull(ListPrice,0) as ListPrice, " & vMargin & " from Products p" & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join ProductPacking SP on SP.packingid = P.SalePackingID and SP.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 " & vShowBetween & IIf(CmbGroup.ListIndex > 0, " and groupid ='" & GetGroupID(CmbGroup) & "'", "") & IIf(CmbSubGroup.ListIndex > 0, " and SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex), "") & IIf(CmbCompany.ListIndex > 0, " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex), "") & " Order by p." & CmbSortBy.Text
  
  With CN.Execute(ssql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("ItemCode").Text = IIf(IsNull(!ItemCode), "", !ItemCode)
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("PurPrice").Value = !PurchasePrice
        Grid.Columns("Packing").Text = !PackingName
        Grid.Columns("Multiplier").Value = !Multiplier
        Grid.Columns("SaleMultiplier").Text = !SaleMultiplier
        Grid.Columns("RetailPrice").Value = !RetailPrice
        Grid.Columns("ListPrice").Value = !ListPrice
        Grid.Columns("WSPrice").Value = !WSPrice
        Grid.Columns("Margin").Value = !Margin
'        If (IsNull(!RetailPrice) Or !RetailPrice = 0) Then
'           Grid.Columns("Margin").Value = 0
'        Else
'           Grid.Columns("Margin").Value = Round((IIf(IsNull(!RetailPrice), 0, !RetailPrice) - IIf(IsNull(!PurPrice), 0, !PurPrice)) * 100 / IIf(IsNull(!RetailPrice) Or !RetailPrice = 0, 1, !RetailPrice), 2)
'        End If
        Grid.Columns("DiscPC").Value = !DiscPC
        Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
        Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), 0, !SaleTaxPer)
        Grid.Columns("PCTCode").Value = IIf(IsNull(!PCTCode), "", !PCTCode)
        Grid.Columns("3rdSchedule").Value = IIf(IsNull(!is3rdScheduleItem), 0, !is3rdScheduleItem)
        Grid.Columns("PurDiscPC").Value = !PurDiscPC
        Grid.Columns("PurDiscPer").Value = IIf(IsNull(!PurDiscPer), 0, !PurDiscPer)
        Grid.Columns("MinStockLimit").Value = IIf(IsNull(!MinStockLimit), 0, !MinStockLimit)
        Grid.Columns("MaxStockLimit").Value = IIf(IsNull(!MaxStockLimit), 0, !MaxStockLimit)
        Grid.Columns("Dead").Value = (!isDeadProduct)
        Grid.Columns("Lock").Value = Abs(!IsLocked)
        Grid.Columns("NoCost").Value = Abs(!IsNoCostProduct)
        Grid.Columns("Raw").Value = Abs(!IsRawProduct)
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
  If Err.Number = 91 Then GoTo Abc
  Grid.Redraw = True
  Me.MousePointer = vbDefault
  Call ShowErrorMessage
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
        
'   vParaSQL = " SELECT distinct p.*, isnull(PackingName,'') as PackingName, isnull(PP.Multiplier,0) as Multiplier, isnull(SP.Multiplier,0) as SaleMultiplier,  isnull(round(pr.amount/(isnull(pr.qtypack,0)+pr.qtyloose),2), p.PurPrice) as PurchasePrice, isnull(ListPrice,0) as ListPrice,  " & vMargin & " from Products p" & vbCrLf _
       + " left outer join ProductBarCodes b on p.productid = b.productid " & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join ProductPacking SP on SP.packingid = P.SalePackingID and SP.productid = p.productid" & vbCrLf _
       + " left outer join (select productid, max(SerialNo) as SerialNo from purchasebody group by productid) m on m.productid = p.productid " & vbCrLf _
       + " left outer join purchasebody pr  on m.productid = pr.productid and pr.SerialNo = m.SerialNo" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 " & IIf(TxtProductID.Text = "", "", " and p.ProductID = " & Val(TxtProductID.Text) & " or b.Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'") & IIf(Trim(TxtItemCode.Text) = "", "", " and ItemCode like '" & TxtItemCode.Text & "%'") & " Order by p." & CmbSortBy.Text
       
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vParaSQL, CN, adOpenStatic, adLockReadOnly
    
   Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrptChangePrice.rpt")
   
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   
   Dim vPrinter() As String
   vPrinter = Split(CmbPrinters.Text, ",")
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   
   RptReportViewer.Report.PrintOut False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbCompany_Click()
   On Error GoTo ErrorHandler
   If CmbCompany.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbDepartment_Click()
   On Error GoTo ErrorHandler
   If cmbDepartment.Visible = False Then Exit Sub
   If ActiveControl.Name <> cmbDepartment.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'----------------------------------
Private Sub CmbGroup_Click()
   On Error GoTo ErrorHandler
   If CmbGroup.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbGroup.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ReQuery()
   On Error GoTo ErrorHandler
   Dim vBm As Variant
   Dim i As Integer
    
   Rs.ReQuery
   Me.MousePointer = vbHourglass
   Grid.Redraw = False
   vSuppressUpdateEvent = True
   vBm = Grid.Bookmark
   Grid.MoveFirst
   For i = 0 To Grid.Rows - 1
      'Rs!PurPrice = Grid.Columns("PurPrice").CellValue(Grid.GetBookmark(i))
      Rs!RetailPrice = Grid.Columns("RetailPrice").CellValue(Grid.GetBookmark(i))
      Rs!WSPrice = Grid.Columns("WSPrice").CellValue(Grid.GetBookmark(i))
      Rs!ListPrice = Grid.Columns("ListPrice").CellValue(Grid.GetBookmark(i))
      Rs!DiscPC = Grid.Columns("DiscPC").CellValue(Grid.GetBookmark(i))
      Rs!DiscPer = Grid.Columns("DiscPer").CellValue(Grid.GetBookmark(i))
      Rs!SaleTaxPer = Grid.Columns("SaleTaxPer").CellValue(Grid.GetBookmark(i))
      Rs!PCTCode = Grid.Columns("PCTCOde").CellValue(Grid.GetBookmark(i))
      Rs!is3rdScheduleItem = Grid.Columns("3rdSchedule").CellValue(Grid.GetBookmark(i))
      Rs!MinStockLimit = Grid.Columns("MinStockLimit").CellValue(Grid.GetBookmark(i))
      Rs!MaxStockLimit = Grid.Columns("MaxStockLimit").CellValue(Grid.GetBookmark(i))
      Rs!isDeadProduct = (Grid.Columns("Dead").CellValue(Grid.GetBookmark(i)))
      Rs!IsLocked = Val(Grid.Columns("Lock").CellValue(Grid.GetBookmark(i)))
      Rs!IsNoCostProduct = Val(Grid.Columns("NoCost").CellValue(Grid.GetBookmark(i)))
      Rs!IsRawProduct = Val(Grid.Columns("Raw").CellValue(Grid.GetBookmark(i)))
      
      Rs.Update
   Next i
   Grid.Bookmark = vBm
   Grid.Redraw = True
   Me.MousePointer = vbDefault
   Rs.UpdateBatch
   MsgBox "Your Entries has been Successfully Updated.", vbOKOnly + vbInformation, "Information"
   Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnMemberDisc_Click()
   FrmMembersDiscount.Show vbModal
End Sub

Private Sub PopulateFilter()
   On Error GoTo ErrorHandler
   'If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
Abc:
  Me.MousePointer = vbHourglass
  ssql = "Select distinct p.ProductID, ProductName, PurPrice, WSPrice, ListPrice, RetailPrice, DiscPC, DiscPer, SaleTaxPer, PCTCOde, is3rdScheduleItem, PurDiscPC, PurDiscPer, MinStockLimit, MaxStockLimit, isDeadProduct, isLocked, IsNoCostProduct, IsRawProduct FROM Products p left outer join ProductBarCodes b on p.productid = b.productid where 1=1 " & IIf(TxtProductID.Text = "", "", " and p.ProductID = " & Val(TxtProductID.Text) & " or Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'") & IIf(Trim(TxtItemCode.Text) = "", "", " and ItemCode like '" & TxtItemCode.Text & "%'")
  Rs.Open ssql, CN, adOpenStatic, adLockBatchOptimistic
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
  
 'IIf(ObjRegistry.ShowDiscPurPrice = True, !PurchasePrice, !PurPrice)
 
  If ObjRegistry.ShowDiscPurPrice = True Then
  ssql = " SELECT distinct p.*, isnull(PackingName,'') as PackingName, isnull(PP.Multiplier,0) as Multiplier, isnull(SP.Multiplier,0) as SaleMultiplier,  isnull(round(pr.amount/(isnull(pr.qtypack,0)+pr.qtyloose),2), p.PurPrice) as PurchasePrice, isnull(ListPrice,0) as ListPrice,  " & vMargin & " from Products p" & vbCrLf _
       + " left outer join ProductBarCodes b on p.productid = b.productid " & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join ProductPacking SP on SP.packingid = P.SalePackingID and SP.productid = p.productid" & vbCrLf _
       + " left outer join (select productid, max(SerialNo) as SerialNo from purchasebody group by productid) m on m.productid = p.productid " & vbCrLf _
       + " left outer join purchasebody pr  on m.productid = pr.productid and pr.SerialNo = m.SerialNo" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 " & IIf(TxtProductID.Text = "", "", " and p.ProductID = " & Val(TxtProductID.Text) & " or b.Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'") & IIf(Trim(TxtItemCode.Text) = "", "", " and ItemCode like '" & TxtItemCode.Text & "%'") & " Order by p." & CmbSortBy.Text
  
  End If
    vParaSQL = ssql
    With CN.Execute(ssql)
      Do Until .EOF
         Grid.AddNew
         Grid.Columns("ID").Text = !Productid
         Grid.Columns("ItemCode").Text = IIf(IsNull(!ItemCode), "", !ItemCode)
         Grid.Columns("Name").Text = !ProductName
         Grid.Columns("Packing").Text = !PackingName
         Grid.Columns("Multiplier").Value = !Multiplier
         Grid.Columns("SaleMultiplier").Value = !SaleMultiplier
         Grid.Columns("PurPrice").Value = IIf(ObjRegistry.ShowDiscPurPrice = True, !PurchasePrice, !PurPrice)
         Grid.Columns("RetailPrice").Value = !RetailPrice
         Grid.Columns("ListPrice").Value = !ListPrice
         Grid.Columns("WSPrice").Value = !WSPrice
         Grid.Columns("Margin").Value = !Margin
'         If (IsNull(!RetailPrice) Or !RetailPrice = 0) Then
'            Grid.Columns("Margin").Value = 0
'         Else
'            Grid.Columns("Margin").Value = Round((IIf(IsNull(!RetailPrice), 0, !RetailPrice) - IIf(IsNull(!PurPrice), 0, !PurPrice)) * 100 / IIf(IsNull(!RetailPrice) Or !RetailPrice = 0, 1, !RetailPrice), 2)
'         End If
         Grid.Columns("DiscPC").Value = !DiscPC
         Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
         Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), 0, !SaleTaxPer)
         Grid.Columns("PCTCode").Value = IIf(IsNull(!PCTCode), "", !PCTCode)
         Grid.Columns("3rdSchedule").Value = (IIf(IsNull(!is3rdScheduleItem), 0, !is3rdScheduleItem))
      
         If ObjRegistry.ShowWholeSaleMargin = True Then
            Grid.Columns("DiscVal").Value = Round(!WSPrice * IIf(IsNull(!DiscPer), 0, !DiscPer) / 100, 2)
         Else
            Grid.Columns("DiscVal").Value = Round(!RetailPrice * IIf(IsNull(!DiscPer), 0, !DiscPer) / 100, 2)
         End If
         Grid.Columns("PurDiscPC").Value = !PurDiscPC
         Grid.Columns("PurDiscPer").Value = IIf(IsNull(!PurDiscPer), 0, !PurDiscPer)
         
         Grid.Columns("MinStockLimit").Value = IIf(IsNull(!MinStockLimit), 0, !MinStockLimit)
         Grid.Columns("MaxStockLimit").Value = IIf(IsNull(!MaxStockLimit), 0, !MaxStockLimit)
         Grid.Columns("Lock").Value = !IsLocked
         Grid.Columns("Dead").Value = (!isDeadProduct)
         Grid.Columns("NoCost").Value = (!IsNoCostProduct)
         Grid.Columns("Raw").Value = (!IsRawProduct)
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
  If Err.Number = 91 Then GoTo Abc
  Grid.Redraw = True
  Me.MousePointer = vbDefault
  Call ShowErrorMessage
End Sub

Private Sub PopulateGrid()
   On Error GoTo ErrorHandler
   If Rs.State = adStateOpen Then
     Rs.CancelBatch
     Rs.Close
   End If
   
   Dim vStrSearch As String, vProductSearch As String, vPCTCode As String
   
   'CmbCompany.ListIndex = 0
   vStrSearch = ""
   If ObjRegistry.isShowDepartment Then vStrSearch = vStrSearch & IIf(cmbDepartment.ListIndex = 0, "", " and DepartmentID =" & cmbDepartment.ItemData(cmbDepartment.ListIndex))
   If ObjRegistry.isShowSubDepartment Then vStrSearch = vStrSearch & IIf(cmbSubDepartment.ListIndex = 0, "", " and SubDepartmentID =" & cmbSubDepartment.ItemData(cmbSubDepartment.ListIndex)) & vbCrLf _
   
   
   Me.MousePointer = vbHourglass
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
      
'      / IIf(Val(Grid.Columns("SaleMultiplier").Value) = 0, 1, Val(Grid.Columns("SaleMultiplier").Value)) - Val(Grid.Columns("PurPrice").Value) / IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value))) * 100 / IIf(Val(Grid.Columns("WSPrice").Value) = 0, 1, Val(Grid.Columns("WSPrice").Value) - Val(Grid.Columns("DiscVal").Value)), 2)
'   Else
'      Grid.Columns("Margin").Value = Round((Val(Grid.Columns("RetailPrice").Value) - Val(Grid.Columns("DiscVal").Value) - Round(Val(Grid.Columns("PurPrice").Value) / IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value)), 2)) * 100 / IIf(Val(Grid.Columns("RetailPrice").Value) = 0, 1, Val(Grid.Columns("RetailPrice").Value) - Val(Grid.Columns("DiscVal").Value)), 2)

'  (  / IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value))) * 100 / IIf(Val(Grid.Columns("WSPrice").Value) = 0, 1, Val(Grid.Columns("WSPrice").Value) - Val(Grid.Columns("DiscVal").Value)), 2)
  
  If ObjRegistry.ShowWholeSaleMargin = True Then
      vMargin = "round(case when p.WSPrice = 0 then 0 else ((isnull(p.WSPrice,0)/isnull(sp.multiplier,1) - (p.PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0))/isnull(pp.multiplier,1))*100/ isnull(p.WSPrice,1)) end,3) as margin"
  Else
      vMargin = "round(case when p.RetailPrice = 0 then 0 else (isnull(p.RetailPrice,0) - ((PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0))/isnull(pp.multiplier,1)))*100/ isnull(p.RetailPrice,1) end,3) as margin"
  End If
   
  vStrSearch = vStrSearch & IIf(CmbGroup.ListIndex > 0, " and groupid = '" & GetGroupID(CmbGroup) & "'", "") & IIf(CmbSubGroup.ListIndex > 0, " and SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex), "") & IIf(CmbCompany.ListIndex > 0, " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex), "") & IIf(ChkShowLocikProduct.Value = 0, " and IsLocked = 0", "")
  
  Select Case ActiveControl.Name
      Case CmbCompany.Name, CmbGroup.Name, CmbSubGroup.Name
         vProductSearch = ""
      Case BtnFilter.Name
      If chkSearchAllProductName.Value = 1 Then
         vStrSearch = ""
      End If
      vProductSearch = IIf(TxtProductID.Text = "", "", " and p.ProductID = " & Val(TxtProductID.Text) & " or b.Code = '" & Val(TxtProductID.Text) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'") & IIf(Trim(TxtItemCode.Text) = "", "", " and ItemCode like '" & TxtItemCode.Text & "%'")
      vProductSearch = vProductSearch & IIf(ChkShowLocikProduct.Value = 0, " and IsLocked = 0", "")
  End Select
   
   vPCTCode = ""
   If ChkShowEmptyPCTCode.Value = 1 Then
      vPCTCode = " and PCTCode is null"
   End If
   
   
   ssql = "Select p.* FROM Products p left outer join ProductBarCodes b on p.productid = b.productid where 1=1 " & vStrSearch & vProductSearch
   Rs.Open ssql, CN, adOpenStatic, adLockBatchOptimistic
  
  ssql = " SELECT distinct p.*, isnull(PackingName,'') as PackingName, isnull(PP.Multiplier,0) as Multiplier, isnull(SP.Multiplier,0) as SaleMultiplier,  PurPrice-isnull(purDiscPC*isnull(sp.multiplier,1),0) as PurchasePrice, isnull(ListPrice,0) as ListPrice,  " & vMargin & " from Products p" & vbCrLf _
       + " left outer join ProductBarCodes b on p.productid = b.productid " & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join ProductPacking SP on SP.packingid = P.SalePackingID and SP.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 " & vStrSearch & vProductSearch & vPCTCode & " Order by p." & CmbSortBy.Text
       
   If vParaSave = True Then
   ssql = " SELECT distinct p.*, isnull(PackingName,'') as PackingName, isnull(PP.Multiplier,0) as Multiplier, isnull(SP.Multiplier,0) as SaleMultiplier,  PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0)  as PurchasePrice, isnull(ListPrice,0) as ListPrice,  " & vMargin & " from Products p" & vbCrLf _
       + " left outer join ProductBarCodes b on p.productid = b.productid " & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join ProductPacking SP on SP.packingid = P.SalePackingID and SP.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1  and p.ProductID in (" & vParaProductID & ")"
   End If
 
   
   If ObjRegistry.ShowDiscPurPrice = True Then
      ssql = " SELECT distinct p.*, isnull(PackingName,'') as PackingName, isnull(PP.Multiplier,0) as Multiplier, isnull(SP.Multiplier,0) as SaleMultiplier,  isnull(round(pr.amount/(isnull(pr.qtypack,0)+pr.qtyloose),2), p.PurPrice) as PurchasePrice, isnull(ListPrice,0) as ListPrice,  " & vMargin & " from Products p" & vbCrLf _
          + " left outer join ProductBarCodes b on p.productid = b.productid " & vbCrLf _
          + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
          + " left outer join ProductPacking SP on SP.packingid = P.SalePackingID and SP.productid = p.productid" & vbCrLf _
          + " left outer join (select productid, max(serialno) as serialno from purchasebody group by productid) m on m.productid = p.productid " & vbCrLf _
          + " left outer join purchasebody pr  on m.productid = pr.productid and pr.serialno = m.serialno" & vbCrLf _
          + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 " & vStrSearch & vProductSearch & vPCTCode & " Order by p." & CmbSortBy.Text
   End If
   
'   Rs.Open "Select * FROM Products where 1=1 " & IIf(CmbGroup.ListIndex > 0, " and groupid ='" & GetGroupID(CmbGroup) & "'", "") & IIf(CmbSubGroup.ListIndex > 0, " and SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex), "") & IIf(CmbCompany.ListIndex > 0, " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex), "") & vStrSearch, cn, adOpenStatic, adLockBatchOptimistic

   vParaSQL = ssql
   With CN.Execute(ssql)
      Do Until .EOF
         Grid.AddNew
         Grid.Columns("ID").Text = !Productid
         Grid.Columns("ItemCode").Text = IIf(IsNull(!ItemCode), "", !ItemCode)
         Grid.Columns("Name").Text = !ProductName
         Grid.Columns("Packing").Text = !PackingName
         Grid.Columns("Multiplier").Value = !Multiplier
         Grid.Columns("SaleMultiplier").Value = !SaleMultiplier
         Grid.Columns("PurPrice").Value = IIf(ObjRegistry.ShowDiscPurPrice = True, !PurchasePrice, !PurPrice)
         Grid.Columns("RetailPrice").Value = !RetailPrice
         Grid.Columns("ListPrice").Value = !ListPrice
         Grid.Columns("WSPrice").Value = !WSPrice
         Grid.Columns("Margin").Value = !Margin
'         If (IsNull(!RetailPrice) Or !RetailPrice = 0) Then
'            Grid.Columns("Margin").Value = 0
'         Else
'            Grid.Columns("Margin").Value = Round((IIf(IsNull(!RetailPrice), 0, !RetailPrice) - IIf(IsNull(!PurPrice), 0, !PurPrice)) * 100 / IIf(IsNull(!RetailPrice) Or !RetailPrice = 0, 1, !RetailPrice), 2)
'         End If
         Grid.Columns("DiscPC").Value = !DiscPC
         Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
         Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), 0, !SaleTaxPer)
         Grid.Columns("PCTCode").Value = IIf(IsNull(!PCTCode), "", !PCTCode)
         Grid.Columns("3rdSchedule").Value = (IIf(IsNull(!is3rdScheduleItem), 0, !is3rdScheduleItem))
      
         If ObjRegistry.ShowWholeSaleMargin = True Then
            Grid.Columns("DiscVal").Value = Round(!WSPrice * IIf(IsNull(!DiscPer), 0, !DiscPer) / 100, 2)
         Else
            Grid.Columns("DiscVal").Value = Round(!RetailPrice * IIf(IsNull(!DiscPer), 0, !DiscPer) / 100, 2)
         End If
         Grid.Columns("PurDiscPC").Value = !PurDiscPC
         Grid.Columns("PurDiscPer").Value = IIf(IsNull(!PurDiscPer), 0, !PurDiscPer)
         
         Grid.Columns("MinStockLimit").Value = IIf(IsNull(!MinStockLimit), 0, !MinStockLimit)
         Grid.Columns("MaxStockLimit").Value = IIf(IsNull(!MaxStockLimit), 0, !MaxStockLimit)
         
         Grid.Columns("Dead").Value = (!isDeadProduct)

         Grid.Columns("Lock").Value = !IsLocked
         Grid.Columns("NoCost").Value = (!IsNoCostProduct)
         Grid.Columns("Raw").Value = (!IsRawProduct)
         Grid.Update
         .MoveNext
      Loop
   End With
   vSuppressUpdateEvent = False
   Grid.Redraw = True
   Grid.MoveFirst
   'If Grid.Visible Then Grid.SetFocus
   Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
  Call CmbGroup_Click
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   
   Call DeleteTempActivityLogBin(vRandomID)
   Grid.Update
   Rs.MoveFirst
   While Not Rs.EOF
      If Rs.EditMode <> adEditNone Then
'         Call ActivityLog("Change Price", eEdit, , , Rs!Productid)
          Rs!IsSync = 0
          Rs!modified_on = Now
      End If
      Rs.MoveNext
   Wend
   Rs.UpdateBatch
   MsgBox "Your Entries has been Successfully Updated.", vbOKOnly + vbInformation, "Information"
   vRandomID = Rnd() * 11111 & " " & Format(Now, "dd/mm hh:mm:ss")
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Function GetGroupID(cmb As ComboBox) As String
    On Error GoTo ErrorHandler
    If cmb.ListIndex < 0 Then Exit Function
    GetGroupID = Chr(Left(cmb.ItemData(cmb.ListIndex), 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 3, 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 5, 2))
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub CmbSortBy_Click()
   On Error GoTo ErrorHandler
   If CmbSortBy.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbSortBy.Name Then Exit Sub
   PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub cmbSubDepartment_Click()
   On Error GoTo ErrorHandler
   If cmbSubDepartment.Visible = False Then Exit Sub
   If ActiveControl.Name <> cmbSubDepartment.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbSubGroup_Change()
   On Error GoTo ErrorHandler
   If CmbSubGroup.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbSubGroup.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbSubGroup_Click()
On Error GoTo ErrorHandler
   If CmbSubGroup.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbSubGroup.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Change Price"
   CmbSortBy.Clear
   CmbSortBy.AddItem "ProductID"
   CmbSortBy.AddItem "ProductName"
   CmbSortBy.AddItem "PCTCode"
   CmbSortBy.AddItem "is3rdScheduleitem"
   CmbGroup.Clear
   CmbSubGroup.Clear
   CmbCompany.Clear
         
   vRandomID = Rnd() * 11111 & " " & Format(Now, "dd/mm hh:mm:ss")
   Grid.Columns("SaleTaxPer").Visible = ObjRegistry.ShowSaleTax
   Grid.Columns("PCTCode").Visible = ObjRegistry.ShowSaleTax
   Grid.Columns("3rdSchedule").Visible = ObjRegistry.ShowSaleTax
   BtnApplySaleTax.Visible = ObjRegistry.ShowSaleTax
   LblSaleTaxPer.Visible = ObjRegistry.ShowSaleTax
   TxtSaleTaxPer.Visible = ObjRegistry.ShowSaleTax
   
   FrmHistory.Visible = ObjRegistry.isShowListPrice
   
   Grid.Columns("PurPrice").Locked = ObjRegistry.LockPurPrice
   Grid.Columns("ListPrice").Visible = ObjRegistry.isShowListPrice
   Grid.Columns("ListPrice").Width = 65
   
   With CN.Execute("Select * FROM Groups order by GroupName")
      CmbGroup.AddItem "All Groups"
      CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left("000", 1)) & Asc(Mid("000", 2, 1)) & Asc(Mid("000", 3, 1))
      Do Until .EOF
         CmbGroup.AddItem !GroupName
         CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
         .MoveNext
      Loop
   End With
   With CN.Execute("Select * FROM SubGroups order by SubGroupName")
      CmbSubGroup.AddItem "All SubGroups"
      CmbSubGroup.ItemData(CmbSubGroup.NewIndex) = 0
      Do Until .EOF
         CmbSubGroup.AddItem !SubGroupName
         CmbSubGroup.ItemData(CmbSubGroup.NewIndex) = !SubGroupID
         .MoveNext
      Loop
   End With
   With CN.Execute("Select * FROM Companies order by CompanyName")
      CmbCompany.AddItem "All Companies"
      CmbCompany.ItemData(CmbCompany.NewIndex) = 0
      Do Until .EOF
         CmbCompany.AddItem !CompanyName
         CmbCompany.ItemData(CmbCompany.NewIndex) = !companyid
         .MoveNext
      Loop
   End With
   If ObjUserSecurity.IsAdministrator Or ObjUserSecurity.IsEditDefination = True Then
    Grid.Columns("Name").Locked = False
   Else
    Grid.Columns("Name").Locked = True
   End If
   ' Item Code visible
   Grid.Columns("ItemCode").Visible = ObjRegistry.ShowColourSize
   TxtItemCode.Visible = ObjRegistry.ShowColourSize
   LblItemCode.Visible = ObjRegistry.ShowColourSize
   cmbDepartment.Visible = ObjRegistry.isShowDepartment
   LblDepartment.Visible = ObjRegistry.isShowDepartment
   cmbSubDepartment.Visible = ObjRegistry.isShowSubDepartment
   LblSubDepartment.Visible = ObjRegistry.isShowSubDepartment
   
   If ObjRegistry.isShowDepartment Then
      cmbDepartment.Clear
      With CN.Execute("Select * FROM Departments Order By Department")
         cmbDepartment.AddItem "All Departments"
         cmbDepartment.ItemData(cmbDepartment.NewIndex) = 0
         Do Until .EOF
            cmbDepartment.AddItem !Department
            cmbDepartment.ItemData(cmbDepartment.NewIndex) = !DepartmentID
            .MoveNext
         Loop
      End With
      cmbDepartment.ListIndex = 0
   End If
   If ObjRegistry.isShowSubDepartment Then
      cmbSubDepartment.Clear
      With CN.Execute("Select * FROM SubDepartments Order By SubDepartmentName")
         cmbSubDepartment.AddItem "All SubDepartments"
         cmbSubDepartment.ItemData(cmbSubDepartment.NewIndex) = 0
         Do Until .EOF
            cmbSubDepartment.AddItem !SubDepartmentName
            cmbSubDepartment.ItemData(cmbSubDepartment.NewIndex) = !SubDepartmentID
            .MoveNext
         Loop
      End With
      cmbSubDepartment.ListIndex = 0
   End If
   
   If ObjRegistry.isShowSubDepartment Or ObjRegistry.isShowDepartment Then
      CmbCompany.ListIndex = 0
   Else
      If CmbCompany.ListCount > 0 Then CmbCompany.ListIndex = 1 Else CmbCompany.ListIndex = 0
   End If
   CmbGroup.ListIndex = 0
   CmbSubGroup.ListIndex = 0
   CmbSortBy.ListIndex = 1
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   
   CmbPrinters.Clear
   CmbPrinters.AddItem "Default,winspool,LPT1"
   Dim p
   For Each p In Printers
      CmbPrinters.AddItem p.DeviceName & "," & p.DriverName & "," & p.Port
   Next p
   CmbPrinters.ListIndex = 0
   If ObjUserSecurity.ChangePriceFormOpenAsLogin = True Then
      CmbCompany.ListIndex = 0
      ShowMarginBelowZero
   End If
'   If vParaSave = True Then Call BtnFilter_Click
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name <> Grid.Name Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
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
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case CmbCompany.Name: If FunSelectCompany() = True Then CmbCompany.SetFocus
         Case CmbGroup.Name: If FunSelectGroup() = True Then CmbGroup.SetFocus
         Case CmbSubGroup.Name: If FunSelectSubGroup() = True Then CmbSubGroup.SetFocus
      End Select
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then BtnSave.Enabled = True
   End If
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

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  If Grid.Columns(ColIndex).Text = "" Then Grid.Columns(ColIndex).Text = "0"
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   On Error GoTo ErrorHandler
   If vSuppressUpdateEvent Then Exit Sub
   Rs.Find "ProductID = " & Val(Grid.Columns("ID").Text), , adSearchForward, 1
   If Rs.EOF Then MsgBox "Cannot Locate Record for updation. Please Try again", vbCritical, "Error": Cancel = True: Exit Sub
   
   If ObjRegistry.SalePriceLessThanPurchase = True Then
      If ObjRegistry.ShowWholeSaleMargin = True Then
         If Val(Grid.Columns("WSPrice").Value) < Val(Grid.Columns("PurPrice").Value) Then
            If ObjUserSecurity.SalePriceMustBeLessThanPurchase = False Then
               If MsgBox("Whole Sale Price ( " & Val(Grid.Columns("WSPrice").Value) & " ) is Less Than Purchase Price ( " & Val(Grid.Columns("PurPrice").Value) & " ). Do you want to change it?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then
                  Grid.Columns("WSPrice").Value = Rs!WSPrice
      '            CalculateMargin
                  Exit Sub
               End If
            Else
               MsgBox "Whole Sale Price ( " & Val(Grid.Columns("WSPrice").Value) & " ) is Less Than Purchase Price ( " & Val(Grid.Columns("PurPrice").Value) & ")", vbInformation + vbOKOnly
               Grid.Columns("WSPrice").Value = Rs!WSPrice
               Exit Sub
            End If
         End If
         If Val(Grid.Columns("ListPrice").Value) < Val(Grid.Columns("PurPrice").Value) And ObjRegistry.isShowListPrice = True Then
            If ObjUserSecurity.SalePriceMustBeLessThanPurchase = False Then
               If MsgBox("List Price ( " & Val(Grid.Columns("ListPrice").Value) & " ) is Less Than Purchase Price ( " & Val(Grid.Columns("PurPrice").Value) & " ). Do you want to change it?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then
                  Grid.Columns("WSPrice").Value = Rs!WSPrice
      '            CalculateMargin
                  Exit Sub
               End If
            Else
               MsgBox "List Price ( " & Val(Grid.Columns("ListPrice").Value) & " ) is Less Than Purchase Price ( " & Val(Grid.Columns("PurPrice").Value) & ")", vbInformation + vbOKOnly
               Grid.Columns("ListPrice").Value = Rs!ListPrice
               Exit Sub
            End If
         End If
         If Val(Grid.Columns("RetailPrice").Value) < Round(Val(Grid.Columns("WSPrice").Value) / IIf(ObjRegistry.isShowListPrice, 1, IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value))), 2) Then
            If ObjUserSecurity.SalePriceMustBeLessThanPurchase = False Then
               If MsgBox("Retial Price ( " & Val(Grid.Columns("RetailPrice").Value) & " ) is Less Than Whole Sale Price ( " & Round(Val(Grid.Columns("WSPrice").Value) / IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value)), 2) & " ). Do you want to change it?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then
                  Grid.Columns("RetailPrice").Value = Rs!RetailPrice
      '            CalculateMargin
                  Exit Sub
               End If
            Else
               MsgBox "Whole Sale Price ( " & Val(Grid.Columns("WSPrice").Value) & " ) is Less Than Purchase Price ( " & Val(Grid.Columns("PurPrice").Value) & ")", vbInformation + vbOKOnly
               Grid.Columns("RetailPrice").Value = Rs!RetailPrice
               Exit Sub
            End If
         End If
      Else
         If Val(Grid.Columns("RetailPrice").Value) < Round(Val(Grid.Columns("PurPrice").Value) / IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value)), 2) Then
            If ObjUserSecurity.SalePriceMustBeLessThanPurchase = False Then
               If MsgBox("Retial Price ( " & Val(Grid.Columns("RetailPrice").Value) & " ) is Less Than Purchase Price ( " & Round(Val(Grid.Columns("PurPrice").Value) / IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value)), 2) & " ). Do you want to change it?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then
                  Grid.Columns("RetailPrice").Value = Rs!RetailPrice
      '            CalculateMargin
                  Exit Sub
               End If
            Else
               
               MsgBox "Retail Price ( " & Val(Grid.Columns("RetailPrice").Value) & " ) is Less Than Purchase Price ( " & Val(Grid.Columns("PurPrice").Value) & ")", vbInformation + vbOKOnly
               Grid.Columns("RetailPrice").Value = Rs!RetailPrice
               Exit Sub
            End If
         End If
      End If
   End If
'   If ObjRegistry.isShowListPrice = True Then
'      If Grid.Columns("MinStockLimit").Value = "" Then
'         MsgBox "Min Stock Limit Must be enter.", vbInformation, "Alert"
'         Exit Sub
'      End If
'   End If
   Call ActivityLogBin("", eFrmChangePrice, eEdit, Grid.Columns("ID").Text, Date, "Effected Code-" & Grid.Columns("ID").Text & " Pur Price-" & Val(Rs!PurPrice) & " WSPrice-" & Val(Rs!WSPrice) & " Retail Price-" & Val(Rs!RetailPrice) & " Pur Disc-" & Val(Rs!PurDiscPC) & " Sale Disc-" & Val(Rs!DiscPC))
   Rs!ProductName = Grid.Columns("Name").Text
   If ObjRegistry.LockPurPrice = False Then Rs!PurPrice = Val(Grid.Columns("PurPrice").Value)
   Rs!WSPrice = Val(Grid.Columns("WSPrice").Value)
   Rs!ListPrice = Val(Grid.Columns("ListPrice").Value)
   Rs!RetailPrice = Val(Grid.Columns("RetailPrice").Value)
   Rs!DiscPC = Val(Grid.Columns("DiscPC").Value)
   Rs!DiscPer = Val(Grid.Columns("DiscPer").Value)
   Rs!SaleTaxPer = Val(Grid.Columns("SaleTaxPer").Value)
   Rs!PCTCode = Grid.Columns("PCTCode").Text
   Rs!is3rdScheduleItem = Val(Grid.Columns("3rdSchedule").Value)
   Rs!PurDiscPC = Val(Grid.Columns("PurDiscPC").Value)
   Rs!PurDiscPer = Val(Grid.Columns("PurDiscPer").Value)
   Rs!MinStockLimit = Val(Grid.Columns("MinStockLimit").Value)
   Rs!MaxStockLimit = Val(Grid.Columns("MaxStockLimit").Value)
   Rs!isDeadProduct = (Grid.Columns("Dead").Value)
   Rs!IsLocked = Val(Grid.Columns("Lock").Value)
   Rs!IsNoCostProduct = Val(Grid.Columns("NoCost").Value)
   Rs!IsRawProduct = Val(Grid.Columns("Raw").Value)
   Rs!is3rdScheduleItem = Val(Grid.Columns("3rdSchedule").Value)
   Call ActivityLogBin("", eFrmChangePrice, eEdit, Grid.Columns("ID").Text, Date, "Updated Code-" & Grid.Columns("ID").Text & " Pur Price-" & Val(Grid.Columns("PurPrice").Text) & " WSPrice-" & Val(Grid.Columns("WSPrice").Text) & " Retail Price-" & Val(Grid.Columns("RetailPrice").Text) & " Pur Disc-" & Val(Grid.Columns("PurDiscPC").Text) & " Sale Disc-" & Val(Grid.Columns("DiscPC").Text))
   Call ActivityLogBin(vRandomID, eFrmStockAdjustment, eAddTempRecord, Grid.Columns("ID").Text, Date, "Pending Update Code-" & Grid.Columns("ID").Text & " Pur Price-" & Val(Grid.Columns("PurPrice").Text) & " WSPrice-" & Val(Grid.Columns("WSPrice").Text) & " Retail Price-" & Val(Grid.Columns("RetailPrice").Text) & " Pur Disc-" & Val(Grid.Columns("PurDiscPC").Text) & " Sale Disc-" & Val(Grid.Columns("DiscPC").Text))
   Rs.Update
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_Change()
'   On Error GoTo ErrorHandler
'
'   ' 4 =  "Multiplier"
'   ' 5 =  "PurPrice"
'   ' 6 =  "WSPrice"
'   ' 7 =  "ListPrice"
'   ' 8 =  "RetailPrice"
'   ' 9 =  "Margin"
'   ' 10 = "DiscPC"
'   ' 11 = "DiscPer"
'   ' 19 = "SaleMultiplier"
'
   If Grid.Col = 5 Then 'Pur Price
      If ObjRegistry.ShowWholeSaleMargin = True Then
'         Grid.Columns("Margin").Value = Round((Val(Grid.Columns("WSPrice").Value) / IIf(Val(Grid.Columns("SaleMultiplier").Value) = 0, 1, Val(Grid.Columns("SaleMultiplier").Value)) - Val(Grid.Columns("PurPrice").Value) / IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value))) * 100 / IIf(Val(Grid.Columns("WSPrice").Value) = 0, 1, Val(Grid.Columns("WSPrice").Value)), 2)
         Grid.Columns("DiscPC").Value = Round((Val(Grid.Columns("WSPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
         Grid.Columns("DiscPer").Value = Round((Val(Grid.Columns("DiscPC").Value) * 100) / IIf(Val(Grid.Columns("WSPrice").Value) = 0, 1, Val(Grid.Columns("WSPrice").Value)), 2)
         Grid.Columns("DiscVal").Value = Round(Val(Grid.Columns("WSPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100, 2)
         Call CalculateMargin
         If Val(Grid.Columns("WSPrice").Value) = 0 Then Grid.Columns("Margin").Value = 0
       Else
'         Grid.Columns("Margin").Value = Round((Val(Grid.Columns("RetailPrice").Value) - Val(Grid.Columns("PurPrice").Value)) * 100 / IIf(Val(Grid.Columns("RetailPrice").Value) = 0, 1, Val(Grid.Columns("RetailPrice").Value)), 2)
         Grid.Columns("DiscPC").Value = Round((Val(Grid.Columns("RetailPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
         Grid.Columns("DiscPer").Value = Round((Val(Grid.Columns("DiscPC").Value) * 100) / IIf(Val(Grid.Columns("RetailPrice").Value) = 0, 1, Val(Grid.Columns("RetailPrice").Value)), 2)
         Grid.Columns("DiscVal").Value = Round(Val(Grid.Columns("RetailPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100, 2)
         Call CalculateMargin
       End If

   End If

   If Grid.Col = 7 Then 'T Price
'       Grid.Columns("Margin").Value = Round((Val(Grid.Columns("WSPrice").Value) / IIf(Val(Grid.Columns("SaleMultiplier").Value) = 0, 1, Val(Grid.Columns("SaleMultiplier").Value)) - Val(Grid.Columns("PurPrice").Value) / IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value))) * 100 / IIf(Val(Grid.Columns("WSPrice").Value) = 0, 1, Val(Grid.Columns("WSPrice").Value)), 2)
       Grid.Columns("DiscPC").Value = Round((Val(Grid.Columns("WSPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
       Grid.Columns("DiscPer").Value = Round((Val(Grid.Columns("DiscPC").Value) * 100) / IIf(Val(Grid.Columns("WSPrice").Value) = 0, 1, Val(Grid.Columns("WSPrice").Value)), 2)
       Grid.Columns("DiscVal").Value = Round(Val(Grid.Columns("WSPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100, 2)
'       Grid.Columns("Margin").Value = Round(((Val(Grid.Columns("WSPrice").Value) - Val(Grid.Columns("DiscVal").Value)) / IIf(Val(Grid.Columns("SaleMultiplier").Value) = 0, 1, Val(Grid.Columns("SaleMultiplier").Value)) - Val(Grid.Columns("PurPrice").Value) / IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value))) * 100 / IIf(Val(Grid.Columns("WSPrice").Value) = 0, 1, Val(Grid.Columns("WSPrice").Value) - Val(Grid.Columns("DiscVal").Value)), 2)
       Call CalculateMargin
       If Val(Grid.Columns("WSPrice").Value) = 0 Then Grid.Columns("Margin").Value = 0
   End If

   If Grid.Col = 8 Then 'Retail Price
'       Grid.Columns("Margin").Value = Round((Val(Grid.Columns("RetailPrice").Value) - (Val(Grid.Columns("PurPrice").Value / IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value))))) * 100 / IIf(Val(Grid.Columns("RetailPrice").Value) = 0, 1, Val(Grid.Columns("RetailPrice").Value)), 2)
       Grid.Columns("DiscPC").Value = Round((Val(Grid.Columns("RetailPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
       Grid.Columns("DiscPer").Value = Round((Val(Grid.Columns("DiscPC").Value) * 100) / IIf(Val(Grid.Columns("RetailPrice").Value) = 0, 1, Val(Grid.Columns("RetailPrice").Value)), 2)
         Grid.Columns("DiscVal").Value = Round(Val(Grid.Columns("RetailPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100, 2)
       Call CalculateMargin
   End If
'
   If Grid.Col = 9 Then 'Margin
         If Val(Grid.Columns("Margin").Value) = 100 Then Exit Sub
'       Grid.Columns("Margin").Value = Round((Val(Grid.Columns("RetailPrice").Value) - (Val(Grid.Columns("PurPrice").Value / IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value))))) * 100 / IIf(Val(Grid.Columns("RetailPrice").Value) = 0, 1, Val(Grid.Columns("RetailPrice").Value)), 2)
       Grid.Columns("RetailPrice").Value = Round((Val(Grid.Columns("PurPrice").Value) / (1 - (Val(Grid.Columns("Margin").Value) / 100))), 0)
       'Call CalculateMargin
   End If
   
   
   If Grid.Col = 10 Then 'DiscPc
       If ObjRegistry.ShowWholeSaleMargin = True Then
         Grid.Columns("DiscPer").Value = Round((Val(Grid.Columns("DiscPC").Value) * 100) / IIf(Val(Grid.Columns("WSPrice").Value) = 0, 1, Val(Grid.Columns("WSPrice").Value)), 2)
         Grid.Columns("DiscVal").Value = Round(Val(Grid.Columns("WSPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100, 2)
       Else
         Grid.Columns("DiscPer").Value = Round((Val(Grid.Columns("DiscPC").Value) * 100) / IIf(Val(Grid.Columns("RetailPrice").Value) = 0, 1, Val(Grid.Columns("RetailPrice").Value)), 2)
         Grid.Columns("DiscVal").Value = Round(Val(Grid.Columns("RetailPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100, 2)
       End If
       Call CalculateMargin
   End If

   If Grid.Col = 11 Then 'DiscPer
       If ObjRegistry.ShowWholeSaleMargin = True Then
         Grid.Columns("DiscPC").Value = Round((Val(Grid.Columns("WSPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
         Grid.Columns("DiscVal").Value = Round(Val(Grid.Columns("WSPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100, 2)
       Else
         Grid.Columns("DiscPC").Value = Round((Val(Grid.Columns("RetailPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
         Grid.Columns("DiscVal").Value = Round(Val(Grid.Columns("RetailPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100, 2)
      End If
       Call CalculateMargin
   End If

   If Grid.Col = 12 Then 'DiscVal
       If ObjRegistry.ShowWholeSaleMargin = True Then
         Grid.Columns("DiscPC").Value = Round((Val(Grid.Columns("RetailPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
         Grid.Columns("DiscPer").Value = Round((Val(Grid.Columns("DiscPC").Value) * 100) / Val(Grid.Columns("WSPrice").Value), 2)

       Else
         Grid.Columns("DiscPer").Value = Round((Val(Val(Grid.Columns("DiscVal").Value)) * 100) / IIf(Val(Grid.Columns("RetailPrice").Value) = 0, 1, Val(Grid.Columns("RetailPrice").Value)), 3)
         Grid.Columns("DiscPC").Value = Round((Val(Grid.Columns("RetailPrice").Value) * Val(Grid.Columns("DiscPer").Value) / 100), 2)
      End If
      Call CalculateMargin
   End If
'    MsgBox Grid.Columns(15).Name
'   MsgBox Grid.Col
'
'
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
End Sub

Private Sub CalculateMargin()
   On Error GoTo ErrorHandler
   If ObjRegistry.ShowWholeSaleMargin = True Then
      Grid.Columns("Margin").Value = Round(((Val(Grid.Columns("WSPrice").Value) - Val(Grid.Columns("DiscVal").Value)) / IIf(Val(Grid.Columns("SaleMultiplier").Value) = 0, 1, Val(Grid.Columns("SaleMultiplier").Value)) - Val(Grid.Columns("PurPrice").Value) / IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value))) * 100 / IIf(Val(Grid.Columns("WSPrice").Value) = 0, 1, Val(Grid.Columns("WSPrice").Value) - Val(Grid.Columns("DiscVal").Value)), 3)
   Else
      Grid.Columns("Margin").Value = Round((Val(Grid.Columns("RetailPrice").Value) - Val(Grid.Columns("DiscVal").Value) - Round(Val(Grid.Columns("PurPrice").Value) / IIf(Val(Grid.Columns("Multiplier").Value) = 0, 1, Val(Grid.Columns("Multiplier").Value)), 2)) * 100 / IIf(Val(Grid.Columns("RetailPrice").Value) = 0, 1, Val(Grid.Columns("RetailPrice").Value) - Val(Grid.Columns("DiscVal").Value)), 3)
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_GotFocus()
   Grid.Row = 0
   Grid.Col = 0
'   SendKeys "{Right}"
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event vbKeyRight, 1, 1, 1
      KeyCode = 0
   End If
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If ObjRegistry.isShowListPrice Then PopulateDataToHistoryGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub PopulateDataToHistoryGrid()
   On Error GoTo ErrorHandler
   Dim PSQL As String, RSQL As String
   
   vSQL = "select top 1 isnull(b.Price,WSPrice) as Price, isnull(h.BillDate,'01-01-2010') as BillDate from Products p left outer join Salebody b on p.ProductID = b.ProductID" & vbCrLf & _
     " left outer join SaleHeader h  on h.SID = b.SID " & vbCrLf & _
     " where credit = 1 and p.productid = " & Val(Grid.Columns("ID").Text) & " and isnull(h.BillDate ,'01-01-2010') < Getdate() order by h.BillDate Desc"
   
   RSQL = "select top 1 isnull(b.Price,p.RetailPrice) as Price, isnull(h.BillDate ,'01-01-2010') as BillDate from Products p left outer join Salebody b on p.ProductID = b.ProductID" & vbCrLf & _
     " left outer join SaleHeader h  on h.SID = b.SID " & vbCrLf & _
     " where cash = 1 and p.productid = " & Val(Grid.Columns("ID").Text) & " and isnull(h.BillDate ,'01-01-2010') < Getdate() order by h.BillDate Desc"
    
 
   GridHistory.Redraw = False
   GridHistory.MoveFirst
   GridHistory.RemoveAll
   GridHistory.AllowAddNew = True
   GridHistory.AddNew
   With CN.Execute(vSQL)
      If .RecordCount > 0 Then
         GridHistory.Columns("WSPrice").Value = !Price
         PSQL = "select top 1 isnull(b.Price,PurPrice) as Price from Products p left outer join Purchasebody b on p.ProductID = b.ProductID" & vbCrLf & _
           " left outer join PurchaseHeader h  on h.PurID = b.PurID and H.PurchaseDate = b.PurchaseDate " & vbCrLf & _
           " where p.productid = " & Val(Grid.Columns("ID").Text) & " and isnull(h.PurchaseDate,'01-01-2010')  <= '" & !BillDate & "' order by b.PurchaseDate Desc"
         With CN.Execute(PSQL)
            If .RecordCount > 0 Then
               GridHistory.Columns("PurPrice").Value = !Price
            Else
               GridHistory.Columns("PurPrice").Value = 0
            End If
            .Close
         End With
      End If
      GridHistory.Columns("Margin").Value = Val(GridHistory.Columns("WSPrice").Value) - Val(GridHistory.Columns("PurPrice").Value)
      If Val(GridHistory.Columns("WSPrice").Value) <> 0 Then
         GridHistory.Columns("MarginPer").Value = Round((Val(GridHistory.Columns("WSPrice").Value) - Val(GridHistory.Columns("PurPrice").Value)) / GridHistory.Columns("WSPrice").Value * 100, 2)
      End If
      .Close
   End With
   GridHistory.Update
   ' Second Row
   GridHistory.AddNew
   With CN.Execute(RSQL)
      If .RecordCount > 0 Then
         GridHistory.Columns("RetailPrice").Value = !Price
         PSQL = "select top 1 isnull(b.Price,PurPrice) as Price from Products p left outer join Purchasebody b on p.ProductID = b.ProductID" & vbCrLf & _
           " left outer join PurchaseHeader h  on h.PurID = b.PurID and H.PurchaseDate = b.PurchaseDate " & vbCrLf & _
           " where p.productid = " & Val(Grid.Columns("ID").Text) & " and isnull(h.PurchaseDate,'01-01-2010')  <= '" & !BillDate & "' order by b.PurchaseDate Desc"
         With CN.Execute(PSQL)
            If .RecordCount > 0 Then
               GridHistory.Columns("PurPrice").Value = !Price
            Else
               GridHistory.Columns("PurPrice").Value = 0
            End If
            .Close
         End With
      End If
      GridHistory.Columns("Margin").Value = Val(GridHistory.Columns("RetailPrice").Value) - Val(GridHistory.Columns("PurPrice").Value)
      If Val(GridHistory.Columns("RetailPrice").Value) <> 0 Then
         GridHistory.Columns("MarginPer").Value = Round((Val(GridHistory.Columns("RetailPrice").Value) - Val(GridHistory.Columns("PurPrice").Value)) / GridHistory.Columns("RetailPrice").Value * 100, 2)
      End If
      .Close
   End With
   GridHistory.Update
   GridHistory.MoveFirst
   GridHistory.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectCompany() As Boolean
   On Error GoTo ErrorHandler
   SchCompany.Show vbModal, Me
   If SchCompany.ParaOutCompanyID = "" Then FunSelectCompany = False: Exit Function
   FunSelectCompany = FindComboIndex(CmbCompany, SchCompany.ParaOutCompanyID)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectSubGroup() As Boolean
   On Error GoTo ErrorHandler
   SchSubGroup.Show vbModal, Me
   If SchSubGroup.ParaOutSubGroupID = "" Then FunSelectSubGroup = False: Exit Function
   FunSelectSubGroup = FindComboIndex(CmbSubGroup, SchSubGroup.ParaOutSubGroupID)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectGroup() As Boolean
   On Error GoTo ErrorHandler
   SchGroup.Show vbModal, Me
   If SchGroup.ParaOutGroupID = "" Then FunSelectGroup = False: Exit Function
   Dim vGroupID As String
   vGroupID = Asc(Left(SchGroup.ParaOutGroupID, 1)) & Asc(Mid(SchGroup.ParaOutGroupID, 2, 1)) & Asc(Mid(SchGroup.ParaOutGroupID, 3, 1))
   FunSelectGroup = FindComboIndex(CmbGroup, vGroupID)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FindComboIndex(ByVal cmb As ComboBox, ByVal result As String) As Boolean
   On Error GoTo ErrorHandler
   Dim i As Integer
   For i = 0 To cmb.ListCount - 1
      If result = cmb.ItemData(i) Then
         cmb.ListIndex = i
         FindComboIndex = True
         Exit Function
      End If
   Next i
   FindComboIndex = False
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub ShowMarginBelowZero()
   On Error GoTo ErrorHandler
   'If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
Abc:
   Dim vStrSQL As String
   vStrSQL = " and round(case when RetailPrice = 0 then 0 else (isnull(RetailPrice,0) - ((PurPrice-isnull(purDiscPC*isnull(multiplier,1),0))/isnull(pp.multiplier,1)))*100/ isnull(RetailPrice,1) end,3) between " & IIf(Trim(TxtFromMargin.Text) = "", "-9999999", Val(TxtFromMargin.Text)) & " and " & IIf(Trim(TxtToMargin.Text) = "", "9999999", Val(TxtToMargin.Text))
   If ObjRegistry.ShowWholeSaleMargin = True Then
'      vMargin = "round(case when WSPrice = 0 then 0 else (isnull(WSPrice,0) - (PurPrice/isnull(sp.multiplier,1)))*100/ isnull(WSPrice,1) end,3) as margin"
'       vMargin = "round(case when WSPrice = 0 then 0 else (isnull(WSPrice,0) - PurPrice)*100/ isnull(WSPrice,1) end,3) as margin"
'       vShowBetween = " and round(case when WSPrice = 0 then 0 else (isnull(WSPrice,0) - PurPrice)*100/ isnull(WSPrice,1) end,3) between " & IIf(Trim(TxtFromMargin.Text) = "", "-9999999", Val(TxtFromMargin.Text)) & " and " & IIf(Trim(TxtToMargin.Text) = "", "9999999", Val(TxtToMargin.Text))
      vMargin = "round(case when WSPrice = 0 then 0 else (isnull(WSPrice,0)/isnull(sp.multiplier,1) - (PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0))/isnull(pp.multiplier,1))*100/ isnull(WSPrice,1) end,3) as margin"
      vShowBetween = " and round(case when WSPrice = 0 then 0 else (isnull(WSPrice,0)/isnull(sp.multiplier,1) - (PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0))/isnull(pp.multiplier,1))*100/ isnull(WSPrice,1) end,3) between " & IIf(Trim(TxtFromMargin.Text) = "", "-9999999", Val(TxtFromMargin.Text)) & " and " & IIf(Trim(TxtToMargin.Text) = "", "9999999", Val(TxtToMargin.Text))
  Else
      vMargin = "round(case when RetailPrice = 0 then 0 else (isnull(RetailPrice,0) - ((PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0))/isnull(pp.multiplier,1)))*100/ isnull(RetailPrice,1) end,3) as margin"
      vShowBetween = " and round(case when RetailPrice = 0 then 0 else (isnull(RetailPrice,0) - ((PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0))/isnull(pp.multiplier,1)))*100/ isnull(RetailPrice,1) end,3) between " & IIf(Trim(TxtFromMargin.Text) = "", "-9999999", Val(TxtFromMargin.Text)) & " and " & IIf(Trim(TxtToMargin.Text) = "", "9999999", Val(TxtToMargin.Text))
  End If
  If Rs.State = adStateOpen Then
    Rs.CancelBatch
    Rs.Close
  End If
  Me.MousePointer = vbHourglass
    
  ssql = " SELECT distinct p.ProductID, ProductName, PurPrice, WSPrice, RetailPrice, ListPrice, PurDiscPer, PurDiscPC, DiscPC, DiscPer, SaleTaxPer, PCTCOde, is3rdScheduleItem, MinStockLimit, MaxStockLimit, isLocked, IsNoCostProduct, IsRawProduct, isDeadProduct, P.IsSync, modified_on from Products p" & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join ProductPacking SP on SP.packingid = P.SalePackingID and SP.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 " & vShowBetween
  
  Rs.Open ssql, CN, adOpenStatic, adLockBatchOptimistic
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
  
   ssql = " SELECT p.*, isnull(PackingName,'') as PackingName, isnull(pp.Multiplier,0) as Multiplier, isnull(SP.Multiplier,0) as SaleMultiplier, PurPrice-isnull(purDiscPC*isnull(pp.multiplier,1),0) as PurchasePrice, isnull(ListPrice,0) as ListPrice, " & vMargin & " from Products p" & vbCrLf _
       + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
       + " left outer join ProductPacking SP on SP.packingid = P.SalePackingID and SP.productid = p.productid" & vbCrLf _
       + " left outer join Packings pa on pa.PackingID = pp.PackingId where 1=1 " & vShowBetween & IIf(CmbGroup.ListIndex > 0, " and groupid ='" & GetGroupID(CmbGroup) & "'", "") & IIf(CmbSubGroup.ListIndex > 0, " and SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex), "") & IIf(CmbCompany.ListIndex > 0, " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex), "") & " Order by p." & CmbSortBy.Text
  
  With CN.Execute(ssql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("ItemCode").Text = IIf(IsNull(!ItemCode), "", !ItemCode)
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("PurPrice").Value = !PurchasePrice
        Grid.Columns("Packing").Text = !PackingName
        Grid.Columns("Multiplier").Value = !Multiplier
        Grid.Columns("SaleMultiplier").Text = !SaleMultiplier
        Grid.Columns("RetailPrice").Value = !RetailPrice
        Grid.Columns("ListPrice").Value = !ListPrice
        Grid.Columns("WSPrice").Value = !WSPrice
        Grid.Columns("Margin").Value = !Margin
'        If (IsNull(!RetailPrice) Or !RetailPrice = 0) Then
'           Grid.Columns("Margin").Value = 0
'        Else
'           Grid.Columns("Margin").Value = Round((IIf(IsNull(!RetailPrice), 0, !RetailPrice) - IIf(IsNull(!PurPrice), 0, !PurPrice)) * 100 / IIf(IsNull(!RetailPrice) Or !RetailPrice = 0, 1, !RetailPrice), 2)
'        End If
        Grid.Columns("DiscPC").Value = !DiscPC
        Grid.Columns("DiscPer").Value = IIf(IsNull(!DiscPer), 0, !DiscPer)
        Grid.Columns("SaleTaxPer").Value = IIf(IsNull(!SaleTaxPer), 0, !SaleTaxPer)
        
        Grid.Columns("PCTCode").Value = IIf(IsNull(!PCTCode), "", !PCTCode)
        Grid.Columns("3rdSchedule").Value = IIf(IsNull(!is3rdScheduleItem), 0, !is3rdScheduleItem)
        
        Grid.Columns("PurDiscPC").Value = !PurDiscPC
        Grid.Columns("PurDiscPer").Value = IIf(IsNull(!PurDiscPer), 0, !PurDiscPer)
        Grid.Columns("MinStockLimit").Value = IIf(IsNull(!MinStockLimit), 0, !MinStockLimit)
        Grid.Columns("MaxStockLimit").Value = IIf(IsNull(!MaxStockLimit), 0, !MaxStockLimit)
        Grid.Columns("Dead").Value = (!isDeadProduct)

        Grid.Columns("Lock").Value = Abs(!IsLocked)
        Grid.Columns("NoCost").Value = Abs(!IsNoCostProduct)
        Grid.Columns("Raw").Value = Abs(!IsRawProduct)
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
  If Err.Number = 91 Then GoTo Abc
  Grid.Redraw = True
  Me.MousePointer = vbDefault
  Call ShowErrorMessage
End Sub


