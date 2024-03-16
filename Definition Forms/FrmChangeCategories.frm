VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmChangeCategories 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmChangeCategories.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkOrganization 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2528
      TabIndex        =   55
      Top             =   8813
      Width           =   195
   End
   Begin VB.ComboBox cmbOrganization 
      Height          =   315
      Left            =   7463
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   1658
      Width           =   1755
   End
   Begin VB.CheckBox ChkBrand 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   8468
      TabIndex        =   11
      Top             =   7868
      Width           =   195
   End
   Begin VB.ComboBox CmbBrand 
      Height          =   315
      Left            =   9278
      Style           =   2  'Dropdown List
      TabIndex        =   48
      Top             =   1658
      Width           =   1755
   End
   Begin VB.ComboBox CmbDepartment 
      Height          =   315
      Left            =   11078
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   1658
      Width           =   1710
   End
   Begin VB.CheckBox ChkDepartment 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2528
      TabIndex        =   13
      Top             =   8318
      Width           =   195
   End
   Begin VB.CheckBox ChkStore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   8468
      TabIndex        =   15
      Top             =   8363
      Width           =   195
   End
   Begin VB.ComboBox CmbStore 
      Height          =   315
      Left            =   11078
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2288
      Width           =   1710
   End
   Begin VB.CheckBox ChkSubGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2528
      TabIndex        =   9
      Top             =   7823
      Width           =   195
   End
   Begin VB.CheckBox ChkGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   8468
      TabIndex        =   7
      Top             =   7388
      Width           =   195
   End
   Begin VB.CheckBox ChkCompany 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2528
      TabIndex        =   5
      Top             =   7343
      Width           =   195
   End
   Begin VB.CheckBox ChkAll 
      Height          =   225
      Left            =   13013
      TabIndex        =   29
      Top             =   2783
      Width           =   195
   End
   Begin VB.ComboBox CmbSubGroup 
      Height          =   315
      Left            =   9278
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2288
      Width           =   1755
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   5723
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2288
      Width           =   1710
   End
   Begin VB.TextBox TxtProductID 
      Height          =   345
      Left            =   1883
      TabIndex        =   25
      Top             =   2303
      Width           =   975
   End
   Begin JeweledBut.JeweledButton BtnFilter 
      Height          =   315
      Left            =   4688
      TabIndex        =   24
      Top             =   2288
      Width           =   900
      _ExtentX        =   1588
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
      MICON           =   "FrmChangeCategories.frx":0ECA
      BC              =   12632256
      FC              =   0
   End
   Begin VB.TextBox TxtProductName 
      Height          =   345
      Left            =   2873
      TabIndex        =   22
      Top             =   2303
      Width           =   1755
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   7463
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2288
      Width           =   1755
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4545
      Left            =   1868
      TabIndex        =   4
      Top             =   2738
      Width           =   11625
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   10
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
      stylesets(0).Picture=   "FrmChangeCategories.frx":0EE6
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
      stylesets(1).Picture=   "FrmChangeCategories.frx":0F02
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
      Columns.Count   =   10
      Columns(0).Width=   1614
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   5953
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2646
      Columns(2).Caption=   "Company"
      Columns(2).Name =   "Company"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).NumberFormat=   "########.##"
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   2646
      Columns(3).Caption=   "Group"
      Columns(3).Name =   "Group"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   2646
      Columns(4).Caption=   "Sub Group"
      Columns(4).Name =   "SubGroup"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   2646
      Columns(5).Caption=   "Brand"
      Columns(5).Name =   "Brand"
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1323
      Columns(6).Caption=   "Select"
      Columns(6).Name =   "Select"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   0
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   11
      Columns(6).FieldLen=   256
      Columns(6).Style=   2
      Columns(7).Width=   2646
      Columns(7).Caption=   "Store"
      Columns(7).Name =   "Store"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Caption=   "Department"
      Columns(8).Name =   "Department"
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Caption=   "Organization"
      Columns(9).Name =   "Organization"
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20505
      _ExtentY        =   8017
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
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6008
      TabIndex        =   17
      Top             =   9278
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
      MICON           =   "FrmChangeCategories.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7313
      TabIndex        =   18
      Top             =   9278
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
      MICON           =   "FrmChangeCategories.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   9938
      TabIndex        =   19
      Top             =   9278
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
      MICON           =   "FrmChangeCategories.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnCompany 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4553
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   7343
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
      MICON           =   "FrmChangeCategories.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   3953
      TabIndex        =   6
      Top             =   7343
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   556
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
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtCompanyName 
      Height          =   315
      Left            =   4913
      TabIndex        =   31
      Tag             =   "nc"
      Top             =   7343
      Width           =   2415
      _ExtentX        =   4260
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
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9968
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   7343
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
      MICON           =   "FrmChangeCategories.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   9368
      TabIndex        =   8
      Top             =   7343
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   556
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
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   10313
      TabIndex        =   34
      Tag             =   "nc"
      Top             =   7343
      Width           =   2415
      _ExtentX        =   4260
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
      Left            =   4553
      TabIndex        =   36
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   7823
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
      MICON           =   "FrmChangeCategories.frx":0FAA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubGroupID 
      Height          =   315
      Left            =   3953
      TabIndex        =   10
      Top             =   7823
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   556
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
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtSubGroupName 
      Height          =   315
      Left            =   4913
      TabIndex        =   37
      Tag             =   "nc"
      Top             =   7823
      Width           =   2415
      _ExtentX        =   4260
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9953
      TabIndex        =   40
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   8318
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
      MICON           =   "FrmChangeCategories.frx":0FC6
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   9353
      TabIndex        =   16
      Top             =   8318
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   556
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
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   10313
      TabIndex        =   41
      Tag             =   "nc"
      Top             =   8318
      Width           =   2415
      _ExtentX        =   4260
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
   Begin JeweledBut.JeweledButton BtnDepartment 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4553
      TabIndex        =   43
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   8273
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
      MICON           =   "FrmChangeCategories.frx":0FE2
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtDepartmentID 
      Height          =   315
      Left            =   3953
      TabIndex        =   14
      Top             =   8273
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   556
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
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtDepartmentName 
      Height          =   315
      Left            =   4913
      TabIndex        =   44
      Tag             =   "nc"
      Top             =   8273
      Width           =   2415
      _ExtentX        =   4260
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
      Left            =   9953
      TabIndex        =   50
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   7823
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
      MICON           =   "FrmChangeCategories.frx":0FFE
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBrandID 
      Height          =   315
      Left            =   9353
      TabIndex        =   12
      Top             =   7823
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   556
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
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtBrandName 
      Height          =   315
      Left            =   10313
      TabIndex        =   51
      Tag             =   "nc"
      Top             =   7823
      Width           =   2415
      _ExtentX        =   4260
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
   Begin JeweledBut.JeweledButton BtnOrganization 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4553
      TabIndex        =   56
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   8768
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
      MICON           =   "FrmChangeCategories.frx":101A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   3953
      TabIndex        =   57
      Top             =   8768
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   556
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
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtOrganizationName 
      Height          =   315
      Left            =   4913
      TabIndex        =   58
      Tag             =   "nc"
      Top             =   8768
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization"
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
      Left            =   2798
      TabIndex        =   59
      Top             =   8813
      Width           =   1080
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
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
      Left            =   7463
      TabIndex        =   54
      Top             =   1433
      Width           =   1620
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brand"
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
      Left            =   8783
      TabIndex        =   52
      Top             =   7868
      Width           =   510
   End
   Begin VB.Label Label10 
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
      Left            =   9278
      TabIndex        =   49
      Top             =   1433
      Width           =   1050
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
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
      Left            =   11078
      TabIndex        =   47
      Top             =   1433
      Width           =   1530
   End
   Begin VB.Label Label8 
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
      Height          =   195
      Left            =   2798
      TabIndex        =   45
      Top             =   8318
      Width           =   990
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Store"
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
      Left            =   8783
      TabIndex        =   42
      Top             =   8363
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   11078
      TabIndex        =   39
      Top             =   2063
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Group"
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
      Left            =   2798
      TabIndex        =   38
      Top             =   7823
      Width           =   915
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
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
      Left            =   8783
      TabIndex        =   35
      Top             =   7388
      Width           =   525
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
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
      Left            =   2798
      TabIndex        =   32
      Top             =   7373
      Width           =   780
   End
   Begin VB.Label Label5 
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
      Left            =   9278
      TabIndex        =   28
      Top             =   2063
      Width           =   1395
   End
   Begin VB.Label Label4 
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
      Left            =   5723
      TabIndex        =   27
      Top             =   2063
      Width           =   1320
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
      Left            =   1883
      TabIndex        =   26
      Top             =   2063
      Width           =   930
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
      Left            =   2873
      TabIndex        =   23
      Top             =   2063
      Width           =   1215
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Categories"
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
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   21
      Top             =   270
      Width           =   2505
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
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
      Left            =   7463
      TabIndex        =   20
      Top             =   2063
      Width           =   1065
   End
End
Attribute VB_Name = "FrmChangeCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vSuppressUpdateEvent As Boolean
Dim ssql As String, i As Long

Private Sub ChkAll_Click()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.Rows - 1
      Grid.Columns("Select").Value = ChkAll.Value
      Grid.MoveNext
   Next i
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub BtnFilter_Click()
  On Error GoTo ErrorHandler
  Me.MousePointer = vbHourglass
   Dim vWords() As String
   Dim vProductName As String
   vWords = Split(TxtProductName.Text, " ")
   vProductName = ""
   For i = 0 To UBound(vWords)
       vProductName = vProductName & " and Productname like '%" & Replace(vWords(i), "'", "''") & "%'"
   Next
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
   ssql = " SELECT ProductID, ProductName, GroupName, CompanyName, SubGroupName, StoreName, Department, BrandName, OrganizationName" & vbCrLf _
      + " FROM Products p Left Outer join Groups g on g.GroupID = p.GroupID" & vbCrLf _
      + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID" & vbCrLf _
      + " left outer join Companies c on c.CompanyID = p.CompanyID" & vbCrLf _
      + " left outer join Brands b on b.BrandID = p.BrandID" & vbCrLf _
      + " left outer join Stores st on st.StoreID = p.StoreID" & vbCrLf _
      + " left outer join Departments d on d.DepartmentID = p.DepartmentID " & vbCrLf _
      + " left outer join Organizations o on o.OrganizationID = p.OrganizationID " & vbCrLf _
      + " where 1=1 " & IIf(TxtProductID.Text = "", "", " and p.isLocked = 0 and p.ProductID = " & Val(TxtProductID.Text)) & vProductName & " Order by p.ProductID --ProductName"
  
  With cn.Execute(ssql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("Group").Value = !GroupName
        Grid.Columns("SubGroup").Value = !SubGroupName
        Grid.Columns("Company").Value = !CompanyName
        Grid.Columns("Store").Value = !StoreName
        Grid.Columns("Department").Value = !Department
        Grid.Columns("Organization").Value = !OrganizationName
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

Private Sub BtnClear_Click()
   ChkCompany.Value = 0
   ChkGroup.Value = 0
   ChkSubGroup.Value = 0
   ChkBrand.Value = 0
   ChkDepartment.Value = 0
   ChkStore.Value = 0
   ChkOrganization.Value = 0
   Call PopulateGrid
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   Dim vSQL As String
   Grid.Update
   If ChkGroup.Value = 1 Or ChkSubGroup.Value = 1 Or ChkCompany.Value = 1 Or ChkBrand.Value = 1 Or ChkStore.Value = 1 Or ChkDepartment.Value = 1 Or ChkOrganization.Value = 1 Then
      ssql = "Update products set IsSync = 0, "
      If ChkCompany.Value = 1 Then
         vSQL = vSQL & " CompanyID = " & IIf(Trim(TxtCompanyID.Text) = "", "Null", TxtCompanyID.Text) & IIf(vSQL = "", "", ",")
      End If
      If ChkGroup.Value = 1 Then
         vSQL = vSQL & IIf(vSQL = "", "", ",") & " GroupID = " & IIf(Trim(TxtGroupID.Text) = "", "Null", "'" & TxtGroupID.Text & "'")
      End If
      If ChkSubGroup.Value = 1 Then
         vSQL = vSQL & IIf(vSQL = "", "", ",") & " SubGroupID = " & IIf(Trim(TxtSubGroupID.Text) = "", "Null", TxtSubGroupID.Text)
      End If
      If ChkBrand.Value = 1 Then
         vSQL = vSQL & IIf(vSQL = "", "", ",") & " BrandID = " & IIf(Trim(TxtBrandID.Text) = "", "Null", TxtBrandID.Text)
      End If
      If ChkDepartment.Value = 1 Then
         vSQL = vSQL & IIf(vSQL = "", "", ",") & " DepartmentID = " & IIf(Trim(TxtDepartmentID.Text) = "", "Null", TxtDepartmentID.Text)
      End If
      If ChkStore.Value = 1 Then
         vSQL = vSQL & IIf(vSQL = "", "", ",") & " StoreID = " & IIf(Trim(TxtStoreID.Text) = "", "Null", TxtStoreID.Text)
      End If
      If ChkOrganization.Value = 1 Then
         vSQL = vSQL & IIf(vSQL = "", "", ",") & " OrganizationID = " & IIf(Trim(TxtOrganizationID.Text) = "", "Null", TxtOrganizationID.Text)
      End If
      
      '& IIf(ChkGroup.Value = 0, "", " GroupID = " & IIf(Trim(TxtGroupID.Text) = "", "Null", "'" & TxtGroupID.Text & "'")) & vbCrLf _
      '& IIf(ChkSubGroup.Value = 0, "", " SubGroupID = " & IIf(Trim(TxtSubGroupID.Text) = "", "Null", TxtSubGroupID.Text)) & vbCrLf _
      '& IIf(ChkCompany.Value = 0, "", " CompanyID = " & IIf(Trim(TxtCompanyID.Text) = "", "Null", TxtCompanyID.Text)) & vbCrLf _

      ssql = ssql & vSQL
      Grid.MoveFirst
      Grid.Redraw = False
      For i = 0 To Grid.Rows - 1
         If Grid.Columns("Select").Value = True Then
            cn.Execute ssql & " Where ProductID = '" & Grid.Columns("ID").Value & "'"
         End If
         Grid.MoveNext
      Next i
      Grid.Redraw = True
   
      MsgBox "Your Entries has been Successfully Updated.", vbOKOnly + vbInformation, "Information"
   Else
      MsgBox "Your must select at least one Category.", vbOKOnly + vbInformation, "Information"
   End If
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub ChkBrand_Click()
   TxtBrandID.Enabled = ChkBrand.Value
   BtnBrand.Enabled = ChkBrand.Value
   TxtBrandName.Enabled = ChkBrand.Value
End Sub

Private Sub ChkCompany_Click()
   TxtCompanyID.Enabled = ChkCompany.Value
   BtnCompany.Enabled = ChkCompany.Value
   TxtCompanyName.Enabled = ChkCompany.Value
End Sub

Private Sub ChkDepartment_Click()
   TxtDepartmentID.Enabled = ChkDepartment.Value
   BtnDepartment.Enabled = ChkDepartment.Value
   TxtDepartmentName.Enabled = ChkDepartment.Value
End Sub

Private Sub ChkGroup_Click()
   TxtGroupID.Enabled = ChkGroup.Value
   BtnGroup.Enabled = ChkGroup.Value
   TxtGroupName.Enabled = ChkGroup.Value
End Sub

Private Sub ChkOrganization_Click()
   TxtOrganizationID.Enabled = ChkOrganization.Value
   BtnOrganization.Enabled = ChkOrganization.Value
   TxtOrganizationName.Enabled = ChkOrganization.Value
End Sub

Private Sub ChkSubGroup_Click()
   TxtSubGroupID.Enabled = ChkSubGroup.Value
   BtnSubGroup.Enabled = ChkSubGroup.Value
   TxtSubGroupName.Enabled = ChkSubGroup.Value
End Sub

Private Sub ChkStore_Click()
   TxtStoreID.Enabled = ChkStore.Value
   BtnStore.Enabled = ChkStore.Value
   TxtStoreName.Enabled = ChkStore.Value
End Sub

Private Sub CmbBrand_Click()
   On Error GoTo ErrorHandler
   If CmbBrand.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbBrand.Name Then Exit Sub
   Call PopulateGrid
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

Private Sub CmbGroup_Click()
   On Error GoTo ErrorHandler
   If CmbGroup.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbGroup.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbOrganization_Click()
   On Error GoTo ErrorHandler
   If CmbOrganization.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbOrganization.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbStore_Click()
   On Error GoTo ErrorHandler
   If CmbStore.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbStore.Name Then Exit Sub
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

Private Function GetGroupID(cmb As ComboBox) As String
    On Error GoTo ErrorHandler
    If cmb.ListIndex <= 0 Then Exit Function
    GetGroupID = Chr(Left(cmb.ItemData(cmb.ListIndex), 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 3, 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 5, 2))
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub PopulateGrid()
   On Error GoTo ErrorHandler
   Me.MousePointer = vbHourglass
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
  
   ssql = " SELECT ProductID, ProductName, GroupName, CompanyName, SubGroupName, BrandName, StoreName, Department, OrganizationName" & vbCrLf _
      + " FROM Products p left outer join Companies c on c.CompanyID = p.CompanyID" & vbCrLf _
      + " Left Outer join Groups g on g.GroupID = p.GroupID" & vbCrLf _
      + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID" & vbCrLf _
      + " left outer join Brands b on b.BrandID = p.BrandID" & vbCrLf _
      + " left outer join Departments d on d.DepartmentID = p.DepartmentID " & vbCrLf _
      + " left outer join Stores st on st.StoreID = p.StoreID" & vbCrLf _
      + " left outer join Organizations o on o.OrganizationID = p.OrganizationID" & vbCrLf _
      + " where p.isLocked = 0 " & IIf(CmbCompany.ListIndex = 0, "", " and p.CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex)) & vbCrLf _
      + IIf(CmbGroup.ListIndex = 0, "", " and p.GroupID ='" & GetGroupID(CmbGroup) & "'") & vbCrLf _
      + IIf(CmbSubGroup.ListIndex = 0, "", " and p.SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex)) & vbCrLf _
      + IIf(CmbBrand.ListIndex = 0, "", " and p.BrandID =" & CmbBrand.ItemData(CmbBrand.ListIndex)) & vbCrLf _
      + IIf(cmbDepartment.ListIndex = 0, "", " and p.DepartmentID =" & cmbDepartment.ItemData(cmbDepartment.ListIndex)) & vbCrLf _
      + IIf(CmbStore.ListIndex = 0, "", " and p.StoreID =" & CmbStore.ItemData(CmbStore.ListIndex)) & vbCrLf _
      + IIf(CmbOrganization.ListIndex = 0, "", " and p.OrganizationID =" & CmbOrganization.ItemData(CmbOrganization.ListIndex)) & vbCrLf _
      + " Order by ProductName"
  
   With cn.Execute(ssql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("Company").Value = !CompanyName
        Grid.Columns("Group").Value = !GroupName
        Grid.Columns("SubGroup").Value = !SubGroupName
        Grid.Columns("Brand").Value = !BrandName
        Grid.Columns("Store").Value = !StoreName
        Grid.Columns("Department").Value = !Department
        Grid.Columns("Organization").Value = !OrganizationName
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

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Change Categories"
   
   ChkAll.ZOrder 0
   ChkCompany_Click
   ChkGroup_Click
   ChkSubGroup_Click
   ChkBrand_Click
   ChkDepartment_Click
   ChkStore_Click
   ChkOrganization_Click
   
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
   
   CmbOrganization.Clear
   With cn.Execute("Select * FROM Organizations Order By OrganizationName")
      CmbOrganization.AddItem "All Organizations"
      CmbOrganization.ItemData(CmbOrganization.NewIndex) = 0
      Do Until .EOF
         CmbOrganization.AddItem !OrganizationName
         CmbOrganization.ItemData(CmbOrganization.NewIndex) = !OrganizationID
         .MoveNext
      Loop
   End With
   
   CmbStore.Clear
   With cn.Execute("Select * FROM Stores Order By StoreName")
      CmbStore.AddItem "All Stores"
      CmbStore.ItemData(CmbStore.NewIndex) = 0
      Do Until .EOF
         CmbStore.AddItem !StoreName
         CmbStore.ItemData(CmbStore.NewIndex) = !StoreID
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
     
   
   
   
   If ObjRegistry.isShowSubDepartment Or ObjRegistry.isShowDepartment Then
      CmbCompany.ListIndex = 0
   Else
      If CmbCompany.ListCount > 0 Then CmbCompany.ListIndex = 1 Else CmbCompany.ListIndex = 0
   End If
   CmbGroup.ListIndex = 0
   CmbSubGroup.ListIndex = 0
   CmbBrand.ListIndex = 0
   cmbDepartment.ListIndex = 0
   CmbStore.ListIndex = 0
   CmbOrganization.ListIndex = 0
   
   PopulateGrid
'  Grid.Columns("Name").Locked = Not ObjUserSecurity.IsAdministrator
'  If CmbCompany.ListCount > 0 Then CmbCompany.ListIndex = 0
'  CmbGroup.ListIndex = 0
   
   'Call BtnFilter_Click
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
         Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then If TxtGroupID.Enabled Then TxtGroupID.SetFocus Else ChkGroup.SetFocus
         Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then If TxtSubGroupID.Enabled Then TxtSubGroupID.SetFocus Else ChkSubGroup.SetFocus
         Case TxtSubGroupID.Name: If FunSelectSubGroup(ssFunctionKey, True) = True Then If TxtBrandID.Enabled Then TxtBrandID.SetFocus Else ChkBrand.SetFocus
         Case TxtBrandID.Name: If FunSelectBrand(ssFunctionKey, True) = True Then If TxtDepartmentID.Enabled Then TxtDepartmentID.SetFocus Else ChkDepartment.SetFocus
         Case TxtDepartmentID.Name: If FunSelectDepartment(ssFunctionKey, True) = True Then If TxtStoreID.Enabled Then TxtStoreID.SetFocus Else ChkStore.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, True) = True Then If TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, True) = True Then If BtnSave.Enabled Then BtnSave.SetFocus
         Case CmbBrand.Name: If FunSelectCmbBrand() = True Then CmbBrand.SetFocus
         Case CmbCompany.Name: If FunSelectCmbCompany() = True Then CmbCompany.SetFocus
         Case CmbGroup.Name: If FunSelectCmbGroup() = True Then CmbGroup.SetFocus
         Case CmbSubGroup.Name: If FunSelectCmbSubGroup() = True Then CmbSubGroup.SetFocus
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

Private Sub ImgExit_Click()
   Unload Me
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
          Exit Function
      Else
          FunSelectOrganization = False
          .Close
          TxtOrganizationID.Text = ""
          TxtOrganizationName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectGroup(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchGroup.Show vbModal, Me
        If SchGroup.ParaOutGroupID = "" Then FunSelectGroup = False: Exit Function
        TxtGroupID.Text = SchGroup.ParaOutGroupID
    End If
    '---------------------------
    TxtGroupID.Text = Right("000" + CStr(Val(TxtGroupID.Text)), 3)
    vStrSQL = " Select * FROM Groups where GroupID='" & TxtGroupID.Text & "'"
    With cn.Execute(vStrSQL)
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

Private Function FunSelectSubGroup(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSubGroup.Show vbModal, Me
        If SchSubGroup.ParaOutSubGroupID = "" Then FunSelectSubGroup = False: Exit Function
        TxtSubGroupID.Text = SchSubGroup.ParaOutSubGroupID
    End If
    '---------------------------
    vStrSQL = " Select * FROM SubGroups where SubGroupID=" & Val(TxtSubGroupID.Text)
    With cn.Execute(vStrSQL)
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
          Exit Function
      Else
          FunSelectStore = False
          .Close
          TxtStoreID.Text = ""
          TxtStoreName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectDepartment(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchDepartment.Show vbModal, Me
        If SchDepartment.ParaOutDepartmentID = "" Then FunSelectDepartment = False: Exit Function
        TxtDepartmentID.Text = SchDepartment.ParaOutDepartmentID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Departments where DepartmentID=" & Val(TxtDepartmentID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtDepartmentName.Text = !Department
          FunSelectDepartment = True
          .Close
          Exit Function
      Else
          FunSelectDepartment = False
          .Close
          TxtDepartmentID.Text = ""
          TxtDepartmentName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectCompany(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCompany.Show vbModal, Me
        If SchCompany.ParaOutCompanyID = "" Then FunSelectCompany = False: Exit Function
        TxtCompanyID.Text = SchCompany.ParaOutCompanyID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Companies where CompanyID=" & Val(TxtCompanyID.Text)
    With cn.Execute(vStrSQL)
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

Private Function FunSelectBrand(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchBrand.Show vbModal, Me
        If SchBrand.ParaOutBrandID = "" Then FunSelectBrand = False: Exit Function
        TxtBrandID.Text = SchBrand.ParaOutBrandID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Brands where BrandID=" & Val(TxtBrandID.Text)
    With cn.Execute(vStrSQL)
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

Private Sub BtnCompany_Click()
   If FunSelectCompany(ssButton, False) = True Then
     If TxtGroupID.Enabled Then TxtGroupID.SetFocus Else ChkGroup.SetFocus
    Else
     If TxtCompanyID.Enabled Then TxtCompanyID.SetFocus
   End If
End Sub

Private Sub BtnDepartment_Click()
   If FunSelectDepartment(ssButton, False) = True Then
     If TxtStoreID.Enabled Then TxtStoreID.SetFocus
   Else
     If TxtDepartmentID.Enabled Then TxtDepartmentID.SetFocus
   End If
End Sub

Private Sub BtnGroup_Click()
   If FunSelectGroup(ssButton, False) = True Then
     If TxtSubGroupID.Enabled Then TxtSubGroupID.SetFocus Else ChkSubGroup.SetFocus
    Else
     If TxtGroupID.Enabled Then TxtGroupID.SetFocus
   End If
End Sub

Private Sub BtnBrand_Click()
   If FunSelectBrand(ssButton, False) = True Then
     If TxtStoreID.Enabled Then TxtStoreID.SetFocus
   Else
     If TxtBrandID.Enabled Then TxtBrandID.SetFocus
   End If
End Sub

Private Sub BtnSubGroup_Click()
   If FunSelectSubGroup(ssButton, False) = True Then
     If TxtStoreID.Enabled Then TxtStoreID.SetFocus
   Else
     If TxtBrandID.Enabled Then TxtBrandID.SetFocus
   End If
End Sub

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      If TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus
   Else
      If TxtStoreID.Enabled Then TxtStoreID.SetFocus
   End If
End Sub

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

Private Sub TxtDepartmentID_Change()
   If TxtDepartmentID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtDepartmentID.Name Then Exit Sub
   If TxtDepartmentName.Text <> "" Then TxtDepartmentName.Text = ""
End Sub

Private Sub TxtDepartmentID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtDepartmentID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtDepartmentID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectDepartment(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectDepartment(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtGroupID_Change()
   If TxtGroupID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtGroupID.Name Then Exit Sub
   If TxtGroupName.Text <> "" Then TxtGroupName.Text = ""
End Sub

Private Sub TxtGroupID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtGroupID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtGroupName.Text <> "" Then Exit Sub
   If TxtGroupID.Text = "" Then Exit Sub
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

Private Sub TxtStoreID_Change()
   If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then TxtStoreName.Text = ""
End Sub

Private Sub TxtStoreID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtStoreID.Text = "" Then Exit Sub
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
      If BtnSave.Enabled Then BtnSave.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Function FunSelectCmbBrand() As Boolean
   On Error GoTo ErrorHandler
   SchBrand.Show vbModal, Me
   If SchBrand.ParaOutBrandID = "" Then FunSelectCmbBrand = False: Exit Function
   FunSelectCmbBrand = FindComboIndex(CmbBrand, SchBrand.ParaOutBrandID)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectCmbCompany() As Boolean
   On Error GoTo ErrorHandler
   SchCompany.Show vbModal, Me
   If SchCompany.ParaOutCompanyID = "" Then FunSelectCmbCompany = False: Exit Function
   FunSelectCmbCompany = FindComboIndex(CmbCompany, SchCompany.ParaOutCompanyID)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectCmbSubGroup() As Boolean
   On Error GoTo ErrorHandler
   SchSubGroup.Show vbModal, Me
   If SchSubGroup.ParaOutSubGroupID = "" Then FunSelectCmbSubGroup = False: Exit Function
   FunSelectCmbSubGroup = FindComboIndex(CmbSubGroup, SchSubGroup.ParaOutSubGroupID)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectCmbGroup() As Boolean
   On Error GoTo ErrorHandler
   SchGroup.Show vbModal, Me
   If SchGroup.ParaOutGroupID = "" Then FunSelectCmbGroup = False: Exit Function
   Dim vGroupID As String
   vGroupID = Asc(Left(SchGroup.ParaOutGroupID, 1)) & Asc(Mid(SchGroup.ParaOutGroupID, 2, 1)) & Asc(Mid(SchGroup.ParaOutGroupID, 3, 1))
   FunSelectCmbGroup = FindComboIndex(CmbGroup, vGroupID)
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
