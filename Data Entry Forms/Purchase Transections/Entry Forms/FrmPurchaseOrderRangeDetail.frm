VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmPurchaseOrderRangeDetail 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmPurchaseOrderRangeDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbSize 
      Height          =   315
      Left            =   9180
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3765
      Width           =   1170
   End
   Begin VB.ComboBox cmbColour 
      Height          =   315
      Left            =   10350
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3765
      Width           =   1170
   End
   Begin VB.ComboBox cmbSeason 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3780
      Width           =   1890
   End
   Begin VB.ComboBox cmbItemDescription 
      Height          =   315
      Left            =   1845
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3780
      Width           =   1890
   End
   Begin VB.ComboBox cmbDescription 
      Height          =   315
      Left            =   9045
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3105
      Width           =   1890
   End
   Begin VB.ComboBox cmbSubDepartment 
      Height          =   315
      Left            =   3735
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3150
      Width           =   1890
   End
   Begin VB.ComboBox cmbDepartment 
      Height          =   315
      Left            =   1845
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3150
      Width           =   1890
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3120
      Width           =   1665
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   7290
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3780
      Width           =   1890
   End
   Begin VB.ComboBox CmbSortBy 
      Height          =   315
      Left            =   11520
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3765
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ComboBox CmbSubGroup 
      Height          =   315
      Left            =   7305
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3120
      Width           =   1665
   End
   Begin VB.ComboBox CmbBrand 
      Height          =   315
      Left            =   3735
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3780
      Width           =   1665
   End
   Begin VB.TextBox TxtItemCode 
      Height          =   345
      Left            =   1905
      TabIndex        =   0
      Top             =   2055
      Width           =   1110
   End
   Begin VB.TextBox TxtProductName 
      Height          =   345
      Left            =   5100
      TabIndex        =   15
      Top             =   1110
      Width           =   2790
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
      Left            =   13080
      TabIndex        =   20
      Top             =   840
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
         TabIndex        =   21
         Tag             =   "NC"
         Text            =   "FrmPurchaseOrderRangeDetail.frx":0ECA
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
         TabIndex        =   22
         Top             =   90
         Width           =   135
      End
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   5745
      TabIndex        =   16
      Top             =   9480
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
      MICON           =   "FrmPurchaseOrderRangeDetail.frx":0F1C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7050
      TabIndex        =   17
      Top             =   9480
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
      MICON           =   "FrmPurchaseOrderRangeDetail.frx":0F38
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8340
      TabIndex        =   18
      Top             =   9480
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
      MICON           =   "FrmPurchaseOrderRangeDetail.frx":0F54
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnFilter 
      Height          =   315
      Left            =   10215
      TabIndex        =   14
      Top             =   2070
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
      MICON           =   "FrmPurchaseOrderRangeDetail.frx":0F70
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnApply 
      Height          =   315
      Left            =   11880
      TabIndex        =   24
      Top             =   2850
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
      MICON           =   "FrmPurchaseOrderRangeDetail.frx":0F8C
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtQtyLoose 
      Height          =   315
      Left            =   11985
      TabIndex        =   25
      Top             =   2445
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
      Height          =   3960
      Left            =   3285
      TabIndex        =   34
      Top             =   4470
      Width           =   8790
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   9
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
      stylesets(0).Picture=   "FrmPurchaseOrderRangeDetail.frx":0FA8
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
      stylesets(1).Picture=   "FrmPurchaseOrderRangeDetail.frx":0FC4
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
      Columns.Count   =   9
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   1852
      Columns(1).Caption=   "Item Code"
      Columns(1).Name =   "ItemCode"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2037
      Columns(2).Caption=   "Colour"
      Columns(2).Name =   "ColourName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1561
      Columns(3).Caption=   "Size"
      Columns(3).Name =   "SizeName"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   5503
      Columns(4).Caption=   "Product Name"
      Columns(4).Name =   "Name"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   1852
      Columns(5).Caption=   "Pur Price"
      Columns(5).Name =   "PurPrice"
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   5
      Columns(5).FieldLen=   256
      Columns(6).Width=   1640
      Columns(6).Caption=   "Qty (L)"
      Columns(6).Name =   "QtyLoose"
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "ColourID"
      Columns(7).Name =   "ColourID"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "SizeID"
      Columns(8).Name =   "SizeID"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   15505
      _ExtentY        =   6985
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
   Begin JeweledBut.JeweledButton BtnItemCode 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3045
      TabIndex        =   35
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2055
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
      MICON           =   "FrmPurchaseOrderRangeDetail.frx":0FE0
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPurPrice 
      Height          =   315
      Left            =   11970
      TabIndex        =   36
      Top             =   1860
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
      Left            =   6060
      TabIndex        =   38
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2085
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
      MICON           =   "FrmPurchaseOrderRangeDetail.frx":0FFC
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtVenderName 
      Height          =   315
      Left            =   6420
      TabIndex        =   39
      Tag             =   "nc"
      Top             =   2085
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
      Left            =   5040
      TabIndex        =   1
      Top             =   2085
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
      Masked          =   1
      IntegralPoint   =   15
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
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
      Left            =   9180
      TabIndex        =   48
      Top             =   3555
      Width           =   375
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Colour"
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
      Left            =   10395
      TabIndex        =   47
      Top             =   3555
      Width           =   555
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Season"
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
      Left            =   5445
      TabIndex        =   46
      Top             =   3555
      Width           =   645
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Description"
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
      Left            =   1845
      TabIndex        =   45
      Top             =   3555
      Width           =   1395
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   9045
      TabIndex        =   44
      Top             =   2835
      Width           =   975
   End
   Begin VB.Label Label12 
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
      Left            =   3735
      TabIndex        =   43
      Top             =   2880
      Width           =   1380
   End
   Begin VB.Label Label11 
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
      Left            =   1845
      TabIndex        =   42
      Top             =   2880
      Width           =   990
   End
   Begin VB.Label Label10 
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
      Left            =   5040
      TabIndex        =   41
      Top             =   1860
      Width           =   870
   End
   Begin VB.Label Label17 
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
      Left            =   6420
      TabIndex        =   40
      Top             =   1860
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      Height          =   1770
      Left            =   11700
      Top             =   1455
      Width           =   1230
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
      Left            =   11925
      TabIndex        =   37
      Top             =   1635
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
      Left            =   3750
      TabIndex        =   33
      Top             =   3555
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
      Left            =   7305
      TabIndex        =   32
      Top             =   2850
      Width           =   1395
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
      Left            =   11940
      TabIndex        =   31
      Top             =   2220
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
      Left            =   11430
      TabIndex        =   30
      Top             =   3555
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1905
      TabIndex        =   29
      Top             =   1815
      Width           =   870
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
      Left            =   5100
      TabIndex        =   28
      Top             =   870
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
      Left            =   7305
      TabIndex        =   27
      Top             =   3555
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
      Left            =   5640
      TabIndex        =   26
      Top             =   2850
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
      Left            =   12780
      TabIndex        =   23
      Top             =   1545
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Detail"
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
      TabIndex        =   19
      Top             =   270
      Width           =   2880
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "FrmPurchaseOrderRangeDetail"
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
   RsTemp.Fields.Append "ProductID", adInteger
   RsTemp.Fields.Append "ItemCode", adVarChar, 9
   RsTemp.Fields.Append "ColourID", adInteger, 5
   RsTemp.Fields.Append "SizeID", adInteger, 5
   RsTemp.Fields.Append "ProductName", adVarChar, 100
   RsTemp.Fields.Append "Price", adDouble
   RsTemp.Fields.Append "QtyLoose", adDouble
   RsTemp.Open
   
   Me.MousePointer = vbHourglass
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
  
   vSQL = " SELECT p.ProductID, ItemCode, ProductName, pc.ColourID, ColourName, ps.SizeID, SizeName, PurPrice" & vbCrLf _
      + " FROM Products p " & vbCrLf _
      + " inner join ProductColours pc on pc.productID = p.ProductID " & " inner join ProductSizes ps on ps.productID = p.ProductID " & vbCrLf _
      + " inner join Colours c on c.ColourID = pc.ColourID " & " inner join Sizes s on s.SizeID = ps.SizeID" & vbCrLf _
      + " left outer join Parties pr on pr.PartyID = p.VendorID1 " & vbCrLf _
      + " Left Outer join Groups g on g.GroupID = p.GroupID" & " left outer join SubGroups sg on sg.SubGroupID = p.SubGroupID" & vbCrLf _
      + " left outer join Companies cp on cp.CompanyID = p.CompanyID" & " left outer join Brands b on b.BrandID = p.BrandID" & vbCrLf _
      + " where p.isLocked = 0 and ItemCode is not Null " & vProductName & vbCrLf _
      + IIf(TxtItemCode.Text = "", "", " and ItemCode = '" & TxtItemCode.Text & "'") & vbCrLf _
      + IIf(TxtVenderID.Text = "", "", " and pr.PartyID = " & Val(TxtVenderID.Text)) & vbCrLf _
      + IIf(cmbDepartment.ListIndex = 0, "", " and p.DepartmentID =" & cmbDepartment.ItemData(cmbDepartment.ListIndex)) & vbCrLf _
      + IIf(cmbSubDepartment.ListIndex = 0, "", " and p.SubDepartmentID =" & cmbSubDepartment.ItemData(cmbSubDepartment.ListIndex)) & vbCrLf _
      + IIf(CmbGroup.ListIndex = 0, "", " and p.GroupID ='" & GetGroupID(CmbGroup) & "'") & vbCrLf _
      + IIf(CmbSubGroup.ListIndex = 0, "", " and p.SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex)) & vbCrLf _
      + IIf(cmbDescription.ListIndex = 0, "", " and p.DescriptionID =" & cmbDescription.ItemData(cmbDescription.ListIndex)) & vbCrLf _
      + IIf(cmbItemDescription.ListIndex = 0, "", " and p.ItemDescriptionID =" & cmbItemDescription.ItemData(cmbItemDescription.ListIndex)) & vbCrLf _
      + IIf(CmbBrand.ListIndex = 0, "", " and p.BrandID =" & CmbBrand.ItemData(CmbBrand.ListIndex)) & vbCrLf _
      + IIf(cmbSeason.ListIndex = 0, "", " and p.SeasonID =" & cmbSeason.ItemData(cmbSeason.ListIndex)) & vbCrLf _
      + IIf(CmbCompany.ListIndex = 0, "", " and p.CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex)) & vbCrLf _
      + " Order by p.ItemCode"
  
   With cn.Execute(vSQL)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("ItemCode").Text = !ItemCode
        Grid.Columns("ColourID").Text = !ColourID
        Grid.Columns("ColourName").Text = !ColourName
        Grid.Columns("SizeID").Text = !SizeID
        Grid.Columns("SizeName").Text = !SizeName
        Grid.Columns("PurPrice").Value = !PurPrice
        Grid.Columns("QtyLoose").Text = ""
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
   RsTemp.Fields.Append "ProductID", adInteger
   RsTemp.Fields.Append "ItemCode", adVarChar, 9
   RsTemp.Fields.Append "ColourID", adInteger, 5
   RsTemp.Fields.Append "SizeID", adInteger, 5
   RsTemp.Fields.Append "ProductName", adVarChar, 100
   RsTemp.Fields.Append "Price", adDouble
   RsTemp.Fields.Append "QtyLoose", adDouble
   RsTemp.Open
   Unload Me
End Sub

Private Sub BtnFilter_Click()
'   On Error GoTo ErrorHandler
'   vProductID = IIf(Val(TxtToProductID.Text) = 0, IIf(Val(TxtItemCode.Text) = 0, "", " and ProductID = '" & TxtItemCode.Text & "'"), IIf(Val(TxtItemCode.Text) = 0, "", " and ProductID Between '" & TxtItemCode.Text & "' and '" & TxtToProductID.Text & "'"))
'   vWords = Split(TxtProductName.Text, " ")
'   vProductName = ""
'   For i = 0 To UBound(vWords)
'       vProductName = vProductName & " and Productname like '%" & Replace(vWords(i), "'", "''") & "%'"
'   Next
   PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnItemCode_Click()
   If FunSelectItemCode(ssButton, True) = True Then
      TxtVenderID.SetFocus
   Else
      TxtItemCode.SetFocus
   End If
End Sub

Private Sub BtnSelect_Click()
   On Error GoTo ErrorHandler
   Unload Me
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
   CmbSortBy.AddItem "ProductID"
   CmbSortBy.AddItem "ProductName"
   
   cmbDepartment.Clear
   With cn.Execute("Select * FROM Departments Order By Department")
      cmbDepartment.AddItem "All Departments"
      cmbDepartment.ItemData(cmbDepartment.NewIndex) = 0
      Do Until .EOF
         cmbDepartment.AddItem !Department
         cmbDepartment.ItemData(cmbDepartment.NewIndex) = !DepartmentID
         .MoveNext
      Loop
   End With
   
   cmbSubDepartment.Clear
   With cn.Execute("Select * FROM SubDepartments Order By SubDepartmentName")
      cmbSubDepartment.AddItem "All SubDepartments"
      cmbSubDepartment.ItemData(cmbSubDepartment.NewIndex) = 0
      Do Until .EOF
         cmbSubDepartment.AddItem !SubDepartmentName
         cmbSubDepartment.ItemData(cmbSubDepartment.NewIndex) = !SubDepartmentID
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
   
   cmbDescription.Clear
   With cn.Execute("Select * FROM Descriptions Order By DescriptionName")
      cmbDescription.AddItem "All Descriptions"
      cmbDescription.ItemData(cmbDescription.NewIndex) = 0
      Do Until .EOF
         cmbDescription.AddItem !DescriptionName
         cmbDescription.ItemData(cmbDescription.NewIndex) = !DescriptionID
         .MoveNext
      Loop
   End With
   
   cmbItemDescription.Clear
   With cn.Execute("Select * FROM ItemDescription Order By ItemDescName")
      cmbItemDescription.AddItem "All ItemDescriptions"
      cmbItemDescription.ItemData(cmbItemDescription.NewIndex) = 0
      Do Until .EOF
         cmbItemDescription.AddItem !ItemDescName
         cmbItemDescription.ItemData(cmbItemDescription.NewIndex) = !ItemDescID
         .MoveNext
      Loop
   End With
   
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
   
   cmbSeason.Clear
   With cn.Execute("Select * FROM Seasons Order By SeasonName")
      cmbSeason.AddItem "All Seasons"
      cmbSeason.ItemData(cmbSeason.NewIndex) = 0
      Do Until .EOF
         cmbSeason.AddItem !SeasonName
         cmbSeason.ItemData(cmbSeason.NewIndex) = !SeasonID
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
   
   cmbColour.Clear
   With cn.Execute("Select * FROM Colours Order By ColourName")
      cmbColour.AddItem "All Colours"
      cmbColour.ItemData(cmbColour.NewIndex) = 0
      Do Until .EOF
         cmbColour.AddItem !ColourName
         cmbColour.ItemData(cmbColour.NewIndex) = !ColourID
         .MoveNext
      Loop
   End With
     
   cmbSize.Clear
   With cn.Execute("Select * FROM Sizes Order By SizeName")
      cmbSize.AddItem "All Sizes"
      cmbSize.ItemData(cmbSize.NewIndex) = 0
      Do Until .EOF
         cmbSize.AddItem !SizeName
         cmbSize.ItemData(cmbSize.NewIndex) = !SizeID
         .MoveNext
      Loop
   End With
   
   cmbDepartment.ListIndex = 0
   cmbSubDepartment.ListIndex = 0
   CmbGroup.ListIndex = 0
   CmbSubGroup.ListIndex = 0
   cmbDescription.ListIndex = 0
   cmbItemDescription.ListIndex = 0
   CmbBrand.ListIndex = 0
   cmbSeason.ListIndex = 0
   CmbCompany.ListIndex = 0
   cmbColour.ListIndex = 0
   cmbSize.ListIndex = 0
   CmbSortBy.ListIndex = 1
   'PopulateGrid

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
         Case TxtItemCode.Name: If FunSelectItemCode(ssFunctionKey, True) = True Then TxtVenderID.SetFocus
         Case TxtVenderID.Name: If FunSelectVender(ssFunctionKey, True) = True Then TxtVenderID.SetFocus
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
   Grid.Row = 0
   Grid.Col = 0
   'SendKeys "{Right}"
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event vbKeyRight, 1, 1, 1
      KeyCode = 0
   End If
End Sub

Private Sub UpdateRs()
   RsTemp.Filter = "ProductID = " & Val(Grid.Columns("ID").Text) & " and ColourID = '" & Grid.Columns("ColourID").Text & "' and SizeID = '" & Grid.Columns("SizeID").Text & "'"
   If RsTemp.RecordCount = 0 And (Val(Grid.Columns("QtyLoose").Value) > 0 And Val(Grid.Columns("PurPrice").Value) > 0) Then
      RsTemp.AddNew
      RsTemp!Productid = Val(Grid.Columns("ID").Text)
      RsTemp!ItemCode = Grid.Columns("ItemCode").Text
      RsTemp!ColourID = Grid.Columns("ColourID").Text
      RsTemp!SizeID = Grid.Columns("SizeID").Text
      RsTemp!ProductName = Grid.Columns("Name").Text
      RsTemp!QtyLoose = Val(Grid.Columns("QtyLoose").Text)
      RsTemp!Price = Val(Grid.Columns("PurPrice").Text)
   ElseIf RsTemp.RecordCount = 1 And Val(Grid.Columns("QtyLoose").Value) = 0 And Val(Grid.Columns("PurPrice").Value) = 0 Then
      RsTemp.Delete
   ElseIf RsTemp.RecordCount = 1 Then
      RsTemp!QtyLoose = Val(Grid.Columns("QtyLoose").Value)
      RsTemp!Price = Val(Grid.Columns("PurPrice").Value)
      RsTemp.Update
  End If
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtItemCode_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If Trim(TxtItemCode.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectItemCode(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectItemCode(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectItemCode(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchItemCode.Show vbModal, Me
        If SchItemCode.ParaOutItemCode = "" Then FunSelectItemCode = False: Exit Function
        TxtItemCode.Text = SchItemCode.ParaOutItemCode
    End If
    '---------------------------
   FunSelectItemCode = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnVender_Click()
   If FunSelectVender(ssButton, False) = True Then
      TxtVenderID.SetFocus
   Else
      TxtVenderID.SetFocus
   End If
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
              " where BarCode = '" & (TxtVenderID.Text) & "' or (c.AccountNo = " & Val(TxtVenderID.Text) & " and (c.AccountNo like '6%') and c.isDetailed = 1 and c.isLocked = 0)"
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
