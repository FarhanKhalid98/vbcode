VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form RptProductPriceList 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptProductPriceList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CHkIncludeLockProducts 
      Caption         =   "Include Locked Products"
      Height          =   195
      Left            =   11205
      TabIndex        =   53
      Top             =   6345
      Width           =   2640
   End
   Begin VB.CheckBox ChkDisc 
      Caption         =   "Only Discount Products"
      Height          =   195
      Left            =   9000
      TabIndex        =   52
      Top             =   6360
      Width           =   2055
   End
   Begin VB.TextBox TxtSectorName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4133
      TabIndex        =   45
      Top             =   8070
      Width           =   3585
   End
   Begin VB.TextBox TxtZoneName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4133
      TabIndex        =   44
      Top             =   7260
      Width           =   3585
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   8933
      TabIndex        =   37
      Top             =   5865
      Width           =   3000
      Begin VB.OptionButton RdoAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All"
         Height          =   255
         Left            =   2340
         TabIndex        =   40
         Top             =   15
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton RdoNotAllocated 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Not Allocated"
         Height          =   255
         Left            =   1035
         TabIndex        =   39
         Top             =   15
         Width           =   1320
      End
      Begin VB.OptionButton RdoAllocated 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allocated"
         Height          =   255
         Left            =   0
         TabIndex        =   38
         Top             =   15
         Width           =   990
      End
   End
   Begin VB.TextBox TxtZoneID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2753
      MaxLength       =   10
      TabIndex        =   5
      Top             =   7260
      Width           =   1020
   End
   Begin VB.TextBox TxtSectorID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2753
      MaxLength       =   10
      TabIndex        =   6
      Top             =   8070
      Width           =   1020
   End
   Begin VB.ComboBox CmbReportType 
      Height          =   315
      Left            =   8648
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5265
      Width           =   3990
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   11258
      TabIndex        =   11
      Top             =   6720
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Close"
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
      MICON           =   "RptProductPriceList.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      Height          =   420
      Left            =   8520
      TabIndex        =   9
      Top             =   6720
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Pre&view"
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
      MICON           =   "RptProductPriceList.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9863
      TabIndex        =   10
      Top             =   6720
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Print"
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
      MICON           =   "RptProductPriceList.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnCustomer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3773
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   8880
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
      MICON           =   "RptProductPriceList.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCustomerID 
      Height          =   315
      Left            =   2753
      TabIndex        =   7
      Top             =   8880
      Width           =   1020
      _ExtentX        =   1799
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
      IntegralPoint   =   3
   End
   Begin JeweledBut.JeweledButton BtnZone 
      Height          =   330
      Left            =   3773
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7260
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
      MICON           =   "RptProductPriceList.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSector 
      Height          =   330
      Left            =   3773
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8070
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
      MICON           =   "RptProductPriceList.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   2753
      TabIndex        =   0
      Top             =   3210
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
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   3773
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3210
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
      MICON           =   "RptProductPriceList.frx":0F72
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3773
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4020
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
      MICON           =   "RptProductPriceList.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   2753
      TabIndex        =   1
      Top             =   4020
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   3
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
   Begin JeweledBut.JeweledButton BtnSubGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3773
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5640
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
      MICON           =   "RptProductPriceList.frx":0FAA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubGroupID 
      Height          =   315
      Left            =   2753
      TabIndex        =   3
      Top             =   5640
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
   Begin JeweledBut.JeweledButton BtnCompany 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3773
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4830
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
      MICON           =   "RptProductPriceList.frx":0FC6
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   2753
      TabIndex        =   2
      Top             =   4830
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   8738
      TabIndex        =   20
      Top             =   2130
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
   Begin JeweledBut.JeweledButton BtnBrand 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3743
      TabIndex        =   41
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   6450
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
      MICON           =   "RptProductPriceList.frx":0FE2
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBrandID 
      Height          =   315
      Left            =   2723
      TabIndex        =   4
      Top             =   6450
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
   Begin SITextBox.Txt TxtCustomerName 
      Height          =   315
      Left            =   4133
      TabIndex        =   46
      Tag             =   "nc"
      Top             =   8880
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
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   4133
      TabIndex        =   47
      Top             =   3210
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
      Masked          =   5
   End
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   4133
      TabIndex        =   48
      Tag             =   "nc"
      Top             =   4020
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
   Begin SITextBox.Txt TxtSubGroupName 
      Height          =   315
      Left            =   4133
      TabIndex        =   49
      Tag             =   "nc"
      Top             =   5640
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
   Begin SITextBox.Txt TxtCompanyName 
      Height          =   315
      Left            =   4133
      TabIndex        =   50
      Tag             =   "nc"
      Top             =   4830
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
   Begin SITextBox.Txt TxtBrandName 
      Height          =   315
      Left            =   4103
      TabIndex        =   51
      Tag             =   "nc"
      Top             =   6450
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
   Begin VB.Label Label15 
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
      Left            =   2723
      TabIndex        =   43
      Top             =   6197
      Width           =   765
   End
   Begin VB.Label Label27 
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
      Left            =   4103
      TabIndex        =   42
      Top             =   6197
      Width           =   1050
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
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
      Left            =   2753
      TabIndex        =   36
      Top             =   8610
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cusotmer Name"
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
      Left            =   4133
      TabIndex        =   35
      Top             =   8610
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zone Name"
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
      Left            =   4133
      TabIndex        =   34
      Top             =   7000
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zone ID"
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
      Left            =   2753
      TabIndex        =   33
      Top             =   7000
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector ID"
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
      Left            =   2753
      TabIndex        =   32
      Top             =   7803
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector Name"
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
      Left            =   4133
      TabIndex        =   31
      Top             =   7803
      Width           =   1110
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
      Left            =   2753
      TabIndex        =   30
      Top             =   2985
      Width           =   1020
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
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
      Height          =   195
      Left            =   4133
      TabIndex        =   29
      Top             =   2985
      Width           =   1215
   End
   Begin VB.Label Label6 
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
      Left            =   2753
      TabIndex        =   28
      Top             =   3788
      Width           =   1020
   End
   Begin VB.Label Label7 
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
      Left            =   4133
      TabIndex        =   27
      Top             =   3788
      Width           =   1065
   End
   Begin VB.Label Label13 
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
      Left            =   2753
      TabIndex        =   26
      Top             =   5394
      Width           =   1155
   End
   Begin VB.Label Label14 
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
      Left            =   2753
      TabIndex        =   25
      Top             =   4591
      Width           =   1020
   End
   Begin VB.Label Label8 
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
      Left            =   4133
      TabIndex        =   24
      Top             =   4591
      Width           =   1320
   End
   Begin VB.Label Label11 
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
      Left            =   4133
      TabIndex        =   23
      Top             =   5394
      Width           =   1455
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Type"
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
      Left            =   8678
      TabIndex        =   22
      Top             =   5040
      Width           =   1065
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   8738
      TabIndex        =   21
      Top             =   1935
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Price List"
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
      Width           =   2325
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove this Row"
      End
   End
End
Attribute VB_Name = "RptProductPriceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim Application1 As New CRAXDRT.Application
Dim Rs As New ADODB.Recordset
Dim sSql As String
Dim vStrSQL As String

Private Sub BtnBrand_Click()
   If FunSelectBrand(ssButton, False) = True Then
      TxtZoneID.SetFocus
   Else
      TxtBrandID.SetFocus
   End If
End Sub

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
    With CN.Execute(vStrSQL)
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
    With CN.Execute(vStrSQL)
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

Private Function FunSelectGroup(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
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
    vStrSQL = " Select * FROM Groups where GroupID='" & TxtGroupID.Text & "'"
    With CN.Execute(vStrSQL)
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
    With CN.Execute(vStrSQL)
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

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtCode.Text) = "" Then Exit Function
    If Len(TxtCode.Text) <= 5 Then
      TxtCode.Text = Right("00000" + CStr(Val(TxtCode.Text)), 5)
    End If
    If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    vStrSQL = " SELECT p.productid, code, ProductName" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"
  
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         FunSelectProduct = True
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtProductID.Text = ""
         TxtCode.Text = ""
         TxtProductName.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      BtnPreview.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnGroup_Click()
   If FunSelectGroup(ssButton, False) = True Then
      TxtSubGroupID.SetFocus
   Else
      TxtGroupID.SetFocus
   End If
End Sub

Private Sub BtnCompany_Click()
   If FunSelectCompany(ssButton, False) = True Then
      TxtGroupID.SetFocus
   Else
      TxtCompanyID.SetFocus
   End If
End Sub

Private Sub BtnSubGroup_Click()
   If FunSelectSubGroup(ssButton, False) = True Then
      TxtCode.SetFocus
   Else
      TxtSubGroupID.SetFocus
   End If
End Sub

Private Sub CmbReportType_Click()
   If CmbReportType.Text = "Distribution Product Price List Allocated" Then
      Frame1.Visible = True
   Else
      Frame1.Visible = False
   End If
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtCode.Text = ""
      TxtProductName.Text = ""
   End If
End Sub

Private Sub TxtCode_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtCode.Text) = "" Then Exit Sub
   vTemp = Not FunSelectProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectProduct(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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

Private Function FunSelectCustomer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchCustomer.Show vbModal, Me
        If SchCustomer.ParaOutCustomerID = "" Then FunSelectCustomer = False: Exit Function
        TxtCustomerID.Text = SchCustomer.ParaOutCustomerID
    End If
    '---------------------------
    vStrSQL = " SELECT * FROM Parties WHERE PartyType<>'V' And PartyID=" & Val(TxtCustomerID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtCustomerName.Text = !PartyName
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

Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
      CmbReportType.SetFocus
   Else
      TxtCustomerID.SetFocus
   End If
End Sub

Private Sub TxtCustomerID_Change()
   If TxtCustomerID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   If TxtCustomerName.Text <> "" Then TxtCustomerName.Text = ""
End Sub

Private Sub TxtCustomerID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtCustomerID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtCustomerID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectCustomer(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectCustomer(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectZone(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchZone.Show vbModal, Me
        If SchZone.ParaOutZoneID = "" Then FunSelectZone = False: Exit Function
        TxtZoneID.Text = SchZone.ParaOutZoneID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Zones where ZoneID=" & Val(TxtZoneID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtZoneName.Text = !ZoneName
          FunSelectZone = True
          .Close
          Exit Function
      Else
          FunSelectZone = False
          .Close
          TxtZoneID.Text = ""
          TxtZoneName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnZone_Click()
   If FunSelectZone(ssButton, False) = True Then
      TxtSectorID.SetFocus
   Else
      TxtZoneID.SetFocus
   End If
End Sub

Private Sub TxtZoneID_Change()
   If TxtZoneID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtZoneID.Name Then Exit Sub
   If TxtZoneName.Text <> "" Then TxtZoneName.Text = ""
End Sub

Private Sub TxtZoneID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtZoneID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtZoneID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectZone(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectZone(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSector(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSector.Show vbModal, Me
        If SchSector.ParaOutSectorID = "" Then FunSelectSector = False: Exit Function
        TxtSectorID.Text = SchSector.ParaOutSectorID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Sectors where SectorID=" & Val(TxtSectorID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSectorName.Text = !SectorName
          FunSelectSector = True
          .Close
          Exit Function
      Else
          FunSelectSector = False
          .Close
          TxtSectorID.Text = ""
          TxtSectorName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnSector_Click()
   If FunSelectSector(ssButton, False) = True Then
      TxtCustomerID.SetFocus
   Else
      TxtSectorID.SetFocus
   End If
End Sub

Private Sub TxtSectorID_Change()
   If TxtSectorID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSectorID.Name Then Exit Sub
   If TxtSectorName.Text <> "" Then TxtSectorName.Text = ""
End Sub

Private Sub TxtSectorID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSectorID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSectorID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSector(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSector(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
       RptReportViewer.Caption = "Product Price List"
       RptReportViewer.Show vbModal
   End If
End Sub

Private Sub BtnPrint_Click()
    If SetReport Then RptReportViewer.Report.PrintOut False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
         Case vbKeyV
            If BtnPreview.Enabled Then BtnPreview_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtGroupID.SetFocus
         Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then TxtCompanyID.SetFocus
         Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then TxtSubGroupID.SetFocus
         Case TxtSubGroupID.Name: If FunSelectSubGroup(ssFunctionKey, True) = True Then TxtBrandID.SetFocus
         Case TxtBrandID.Name: If FunSelectBrand(ssFunctionKey, True) = True Then TxtZoneID.SetFocus
         Case TxtZoneID.Name: If FunSelectZone(ssFunctionKey, True) = True Then TxtSectorID.SetFocus
         Case TxtSectorID.Name: If FunSelectSector(ssFunctionKey, True) = True Then TxtCustomerID.SetFocus
         Case TxtCustomerID.Name: If FunSelectCustomer(ssFunctionKey, True) = True Then CmbReportType.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Product Price List"
   
   CmbReportType.Clear
   CmbReportType.AddItem "Product Price List"
   CmbReportType.AddItem "Product BarCode List"
   CmbReportType.AddItem "Product Purchase Price List"
   CmbReportType.AddItem "Product Price List (Brand Wise)"
   CmbReportType.AddItem "Product Price List (Customer Wise)"
   CmbReportType.AddItem "Product Price List (Company Wise)"
   CmbReportType.AddItem "Product Price List (Group Wise)"
   CmbReportType.AddItem "Product Price List (Sub Group Wise)"
   CmbReportType.AddItem "Distribution Product Price List Allocated"
   CmbReportType.ListIndex = 0
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   Dim frmObj As Object
   For Each frmObj In Forms
       Set frmObj = Nothing
   Next
   Set RptProductPriceList = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Function SetReport() As Boolean
   On Error GoTo ErrorHandler
   Dim RsReport As New ADODB.Recordset
   SetReport = False
   Me.MousePointer = vbHourglass
   vStrSQL = "Select p.ProductID, ProductName, Desc1, WSPrice, RetailPrice, DiscPC, DiscPer, PurDiscPC, p.GroupID, GroupName, " & vbCrLf _
        + " p.SubGroupID, SubGroupName, p.CompanyID, CompanyName, p.BrandID, BrandName, " & vbCrLf _
        + " Case when isnull(multiplier,0) = 0 then null else WSPrice/multiplier end as WSPriceUnit, " & vbCrLf _
        + " Case when isnull(multiplier,0) = 0 then null else RetailPrice/multiplier end as RetailPriceUnit " & vbCrLf _
        + " from Products p Left Outer Join Packings PK on P.SalePackingID = PK.PackingID " & vbCrLf _
        + " Left Outer Join ProductPacking PP on PP.PackingID = PK.PackingID  and P.ProductId = PP.ProductID" & vbCrLf _
        + " left outer join Companies c on c.CompanyID = p.CompanyID " & vbCrLf _
        + " left outer join Groups g on g.GroupID = p.GroupID " & vbCrLf _
        + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID " & vbCrLf _
        + " left outer join Brands b on b.BrandID = p.BrandID " & vbCrLf _
        + " where 1=1 " & IIf(ChkDisc.Value = 1, " and DiscPc > 0", "") & vbCrLf _
        + IIf(CHkIncludeLockProducts.Value = 0, " and isLocked = 0", "") & vbCrLf _
        + IIf(Trim(TxtCompanyID.Text) = "", "", " And p.CompanyID = " & TxtCompanyID.Text) & vbCrLf _
        + IIf(Trim(TxtGroupID.Text) = "", "", " And p.GroupID = " & TxtGroupID.Text) & vbCrLf _
        + IIf(Trim(TxtSubGroupID.Text) = "", "", " And p.SubGroupID = " & TxtSubGroupID.Text) & vbCrLf _
        + IIf(Trim(TxtBrandID.Text) = "", "", " And p.BrandID = " & TxtBrandID.Text) & vbCrLf _
        + IIf(Trim(TxtCode.Text) = "", "", " And P.ProductID = '" & TxtProductID.Text & "'") & vbCrLf _
        + " Order By ProductName "
   Select Case CmbReportType.Text
   Case "Product Price List"
      vStrSQL = " select p.ProductID, ProductName, CompanyName, GroupName, SubGroupName, BrandName,RetailPrice, PurPrice, 0 Cost, WSPrice, Desc1 " & vbCrLf _
        + " from Products p  " & vbCrLf _
        + " left outer join Companies c on c.CompanyID = p.CompanyID " & vbCrLf _
        + " left outer join Groups g on g.GroupID = p.GroupID " & vbCrLf _
        + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID " & vbCrLf _
        + " left outer join Brands b on b.BrandID = p.BrandID " & vbCrLf _
        + " where 1=1 " & IIf(ChkDisc.Value = 1, " and DiscPc > 0", "") & vbCrLf _
        + IIf(CHkIncludeLockProducts.Value = 0, " and isLocked = 0", "") & vbCrLf _
        + IIf(Trim(TxtCompanyID.Text) = "", "", " And p.CompanyID = " & TxtCompanyID.Text) & vbCrLf _
        + IIf(Trim(TxtGroupID.Text) = "", "", " And p.GroupID = " & TxtGroupID.Text) & vbCrLf _
        + IIf(Trim(TxtSubGroupID.Text) = "", "", " And p.SubGroupID = " & TxtSubGroupID.Text) & vbCrLf _
        + IIf(Trim(TxtBrandID.Text) = "", "", " And p.BrandID = " & TxtBrandID.Text) & vbCrLf _
        + IIf(Trim(TxtCode.Text) = "", "", " And P.ProductID = '" & TxtProductID.Text & "'") & vbCrLf _
        + " Order By ProductName "
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\ListReports\CrptProdcutListNew.rpt")
   Case "Product BarCode List"
      vStrSQL = " select p.ProductID, ProductName, CompanyName, GroupName, SubGroupName, BrandName,RetailPrice, PurPrice, 0 Cost, WSPrice, Desc1, Code, Qty  " & vbCrLf _
        + " from Products p  " & vbCrLf _
        + " Inner join ProductBarcodes PB on PB.ProductID = p.ProductID " & vbCrLf _
        + " left outer join Companies c on c.CompanyID = p.CompanyID " & vbCrLf _
        + " left outer join Groups g on g.GroupID = p.GroupID " & vbCrLf _
        + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID " & vbCrLf _
        + " left outer join Brands b on b.BrandID = p.BrandID " & vbCrLf _
        + " where 1=1 " & IIf(ChkDisc.Value = 1, " and p.DiscPc > 0", "") & vbCrLf _
        + IIf(CHkIncludeLockProducts.Value = 0, " and isLocked = 0", "") & vbCrLf _
        + IIf(Trim(TxtCompanyID.Text) = "", "", " And p.CompanyID = " & TxtCompanyID.Text) & vbCrLf _
        + IIf(Trim(TxtGroupID.Text) = "", "", " And p.GroupID = " & TxtGroupID.Text) & vbCrLf _
        + IIf(Trim(TxtSubGroupID.Text) = "", "", " And p.SubGroupID = " & TxtSubGroupID.Text) & vbCrLf _
        + IIf(Trim(TxtBrandID.Text) = "", "", " And p.BrandID = " & TxtBrandID.Text) & vbCrLf _
        + IIf(Trim(TxtCode.Text) = "", "", " And P.ProductID = '" & TxtProductID.Text & "'") & vbCrLf _
        + " Order By ProductName "
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\ListReports\CrptProdcutListBarCode.rpt")
      'Set RptReportViewer.Report = New CrptProdcutListNew
   Case "Product Purchase Price List"
      vStrSQL = " Select p.ProductID, ProductName, CompanyName, GroupName, SubGroupName, BrandName,RetailPrice, PurPrice, 0 Cost, WSPrice, Desc1," & vbCrLf _
        + " Case when isnull(multiplier,0) = 0 then null else PurPrice/multiplier end as PurPriceUnit, " & vbCrLf _
        + " Case when isnull(multiplier,0) = 0 then null else WSPrice/multiplier end as WSPriceUnit, " & vbCrLf _
        + " Case when isnull(multiplier,0) = 0 then null else RetailPrice/multiplier end as RetailPriceUnit " & vbCrLf _
        + " from Products p Left Outer Join Packings PK on P.PurchasePackingID = PK.PackingID " & vbCrLf _
        + " Left Outer Join ProductPacking PP on PP.PackingID = PK.PackingID  and P.ProductId = PP.ProductID" & vbCrLf _
        + " left outer join Companies c on c.CompanyID = p.CompanyID " & vbCrLf _
        + " left outer join Groups g on g.GroupID = p.GroupID " & vbCrLf _
        + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID " & vbCrLf _
        + " left outer join Brands b on b.BrandID = p.BrandID " & vbCrLf _
        + " where 1=1 " & IIf(ChkDisc.Value = 1, " and p.DiscPc > 0", "") & vbCrLf _
        + IIf(CHkIncludeLockProducts.Value = 0, " and isLocked = 0", "") & vbCrLf _
        + IIf(Trim(TxtCompanyID.Text) = "", "", " And p.CompanyID = " & TxtCompanyID.Text) & vbCrLf _
        + IIf(Trim(TxtGroupID.Text) = "", "", " And p.GroupID = " & TxtGroupID.Text) & vbCrLf _
        + IIf(Trim(TxtSubGroupID.Text) = "", "", " And p.SubGroupID = " & TxtSubGroupID.Text) & vbCrLf _
        + IIf(Trim(TxtBrandID.Text) = "", "", " And p.BrandID = " & TxtBrandID.Text) & vbCrLf _
        + IIf(Trim(TxtCode.Text) = "", "", " And P.ProductID = '" & TxtProductID.Text & "'") & vbCrLf _
        + " Order By ProductName "
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\ListReports\CRptProductPriceList.rpt")
   Case "Product Price List (Brand Wise)"
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\ListReports\CrptDistributionProductPriceListBrandWise.rpt")
   Case "Product Price List (Company Wise)"
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\ListReports\CrptDistributionProductPriceListCompanyWise.rpt")
   Case "Product Price List (Group Wise)"
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\ListReports\CrptDistributionProductPriceListGroupWise.rpt")
   Case "Product Price List (Sub Group Wise)"
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\ListReports\CrptDistributionProductPriceListSubGroupWise.rpt")
   Case "Distribution Product Price List Allocated"
      vStrSQL = "Select p.ProductID, ProductName, Desc1, CompanyName, WSPrice, RetailPrice, p.ZoneID, ZoneName, case when b.ProductID is null then '' else 'Ö' end as Allcated" & vbCrLf _
        + " from (Select * from zones,Products) p " & vbCrLf _
        + " left outer join (Select cp.*,s.* from CustomerProductPrice cp " & vbCrLf _
        + " inner join Parties p on cp.CustomerID = p.PartyID" & vbCrLf _
        + " left outer Join Sectors S on S.SectorID = p.SectorID" & vbCrLf _
        + " left outer join Zones Z on Z.ZoneID = s.ZoneID)b on p.zoneid = b.zoneid and p.productid = b.productid" & vbCrLf _
        + " left outer join Companies c on c.CompanyID = p.CompanyID" & vbCrLf _
        + " left outer join Groups g on g.GroupID = p.GroupID" & vbCrLf _
        + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID" & vbCrLf _
        + " left outer join Brands br on br.BrandID = p.BrandID " & vbCrLf _
        + " where 1=1 " & IIf(ChkDisc.Value = 1, " and P.DiscPc > 0", "") & vbCrLf _
        + IIf(CHkIncludeLockProducts.Value = 0, " and isLocked = 0", "") & vbCrLf _
        + IIf(Trim(TxtCompanyID.Text) = "", "", " And p.CompanyID = " & TxtCompanyID.Text) & vbCrLf _
        + IIf(Trim(TxtGroupID.Text) = "", "", " And p.GroupID = " & TxtGroupID.Text) & vbCrLf _
        + IIf(Trim(TxtSubGroupID.Text) = "", "", " And p.SubGroupID = " & TxtSubGroupID.Text) & vbCrLf _
        + IIf(Trim(TxtBrandID.Text) = "", "", " And p.BrandID = " & TxtBrandID.Text) & vbCrLf _
        + IIf(Trim(TxtCode.Text) = "", "", " And P.ProductID = '" & TxtProductID.Text & "'") & vbCrLf _
        + IIf(Trim(TxtZoneID.Text) = "", "", " And p.ZoneID = " & TxtZoneID.Text) & vbCrLf _
        + IIf(Trim(TxtSectorID.Text) = "", "", " And b.SectorID = " & TxtSectorID.Text) & vbCrLf _
        + IIf(Trim(TxtCustomerID.Text) = "", "", " And b.CustomerID = '" & TxtCustomerID.Text & "'") & vbCrLf _
        + IIf(RdoAll.Value = True, "", IIf(RdoAllocated.Value = True, " and b.ProductID is not null", " and b.ProductID is null")) & vbCrLf _
        + " Order By ProductName "
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\ListReports\CrptDistributionProductPriceListZoneWise.rpt")
   Case "Product Price List (Customer Wise)"
         vStrSQL = " SELECT p.ProductID, ProductName, cp.Price, cp.DiscPer,  P.DiscPC ProdDiscPC, P.DiscPer ProdDiscPer, PurDiscPC, cp.CustomerID, PartyName as CustomerName," & vbCrLf _
            + " p.GroupID, GroupName, p.SubGroupID, SubGroupName, p.CompanyID, CompanyName, st.ZoneID, ZoneName" & vbCrLf _
            + " FROM Products p inner join CustomerProductPrice cp on p.ProductID = cp.ProductID " & vbCrLf _
            + " inner join Parties pt on pt.PartyID = cp.CustomerID" & vbCrLf _
            + " left outer join Groups g on g.groupid = p.groupid" & vbCrLf _
            + " left outer join SubGroups s on s.subgroupid = p.subgroupid" & vbCrLf _
            + " left outer join Brands b on b.BrandID = p.BrandID " & vbCrLf _
            + " left outer join Companies c on c.CompanyID = p.CompanyID " & vbCrLf _
            + " left outer Join Sectors St on St.sectorId = pt.sectorID" & vbCrLf _
            + " left outer join Zones Z on Z.ZoneID = St.ZoneID Where 1=1 " & IIf(ChkDisc.Value = 1, " and P.DiscPc > 0", "") & vbCrLf _
            + IIf(CHkIncludeLockProducts.Value = 0, " and isLocked = 0", "") & vbCrLf _
            + "" & IIf(Trim(TxtZoneID.Text) = "", "", " And z.ZoneID = " & TxtZoneID.Text) & vbCrLf _
            + "" & IIf(Trim(TxtSectorID.Text) = "", "", " And p.SectorID = " & TxtSectorID.Text) & vbCrLf _
            + "" & IIf(Trim(TxtCustomerID.Text) = "", "", " And cp.CustomerID = '" & TxtCustomerID.Text & "'") & vbCrLf _
            + "" & IIf(Trim(TxtCompanyID.Text) = "", "", " And p.CompanyID = " & TxtCompanyID.Text) & vbCrLf _
            + "" & IIf(Trim(TxtGroupID.Text) = "", "", " And p.GroupID = " & TxtGroupID.Text) & vbCrLf _
            + "" & IIf(Trim(TxtSubGroupID.Text) = "", "", " And p.SubGroupID = " & TxtSubGroupID.Text) & vbCrLf _
            + IIf(Trim(TxtBrandID.Text) = "", "", " And p.BrandID = " & TxtBrandID.Text) & vbCrLf _
            + "" & IIf(Trim(TxtCode.Text) = "", "", " And P.ProductID = '" & TxtProductID.Text & "'")
          Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\Reports\ListReports\CrptCustomerProductPrice.rpt")
          RptReportViewer.Report.PaperOrientation = crPortrait
   End Select
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   
   If RsReport.BOF Then
       MsgBox "No record exists.", vbInformation, Me.Caption
       Me.MousePointer = vbDefault
       Exit Function
   End If
   
   RptReportViewer.Report.ReportTitle = CmbReportType.Text
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   RptReportViewer.Report.PaperOrientation = crPortrait
   SetReport = True
   Me.MousePointer = vbDefault
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
