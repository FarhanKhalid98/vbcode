VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmSoftwareDefaultSettings 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnAllow 
      Height          =   285
      Left            =   4650
      TabIndex        =   61
      Top             =   1050
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   503
      TX              =   "Allow"
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
      MICON           =   "FrmSoftwareDefaultSettings.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnShowHide 
      Height          =   285
      Left            =   6690
      TabIndex        =   62
      Top             =   1050
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   503
      TX              =   "Show Hide"
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
      MICON           =   "FrmSoftwareDefaultSettings.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6240
      TabIndex        =   58
      Top             =   10050
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
      MICON           =   "FrmSoftwareDefaultSettings.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7500
      TabIndex        =   59
      Top             =   10050
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
      MICON           =   "FrmSoftwareDefaultSettings.frx":0054
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDefault 
      Height          =   285
      Left            =   2610
      TabIndex        =   60
      Top             =   1050
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   503
      TX              =   "Defaults"
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
      MICON           =   "FrmSoftwareDefaultSettings.frx":0070
      BC              =   14737632
      FC              =   0
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   13185
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin JeweledBut.JeweledButton BtnSMS 
      Height          =   285
      Left            =   8730
      TabIndex        =   76
      Top             =   1050
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   503
      TX              =   "SMS"
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
      MICON           =   "FrmSoftwareDefaultSettings.frx":008C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnEmail 
      Height          =   285
      Left            =   10770
      TabIndex        =   234
      Top             =   1035
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   503
      TX              =   "Email"
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
      MICON           =   "FrmSoftwareDefaultSettings.frx":00A8
      BC              =   14737632
      FC              =   0
   End
   Begin VB.PictureBox PicDefault 
      BorderStyle     =   0  'None
      Height          =   8505
      Left            =   2655
      ScaleHeight     =   8505
      ScaleWidth      =   12300
      TabIndex        =   20
      Top             =   1260
      Width           =   12300
      Begin VB.TextBox TxtChargesName 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   10215
         MaxLength       =   200
         TabIndex        =   225
         Text            =   "Other Charges"
         Top             =   6405
         Width           =   1635
      End
      Begin VB.TextBox TxtRoundfigureInSearchForm 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   9765
         MaxLength       =   2
         TabIndex        =   188
         Text            =   "0"
         Top             =   8190
         Width           =   345
      End
      Begin VB.TextBox TxtGridRowHeight 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1935
         MaxLength       =   2
         TabIndex        =   181
         Text            =   "70"
         Top             =   3780
         Width           =   345
      End
      Begin VB.TextBox TxtAdminClssingFinePerOnShort 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3645
         MaxLength       =   5
         TabIndex        =   178
         Top             =   8190
         Width           =   525
      End
      Begin VB.TextBox TxtCostY 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2610
         MaxLength       =   100
         TabIndex        =   167
         Top             =   4365
         Width           =   435
      End
      Begin VB.TextBox TxtCostX 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2115
         MaxLength       =   100
         TabIndex        =   166
         Top             =   4365
         Width           =   435
      End
      Begin VB.TextBox TxtEmployeeLateRelaxTime 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2700
         MaxLength       =   3
         TabIndex        =   164
         Top             =   4860
         Width           =   345
      End
      Begin VB.TextBox TxtBarcodePrefix 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   162
         Top             =   5220
         Width           =   345
      End
      Begin VB.TextBox TxtPackingChargesPer 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2310
         MaxLength       =   5
         TabIndex        =   158
         Top             =   5610
         Width           =   435
      End
      Begin VB.ComboBox Cmb2ndPrinters 
         Height          =   315
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   146
         Top             =   5415
         Width           =   3315
      End
      Begin VB.CheckBox ChkIsPortrait 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "IsPortrait"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6840
         TabIndex        =   137
         Top             =   4800
         Width           =   1020
      End
      Begin VB.CheckBox ChkIsLegal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "IsLegal"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7920
         TabIndex        =   136
         Top             =   4800
         Width           =   1140
      End
      Begin VB.CheckBox ChkCurrentDateDataEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Current Date Data Entry"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7470
         TabIndex        =   114
         Top             =   5850
         Width           =   2175
      End
      Begin VB.CheckBox ChkIsEntryDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Apply Entry Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7470
         TabIndex        =   100
         Top             =   6165
         Width           =   1500
      End
      Begin VB.TextBox TxtNoofPrints 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3300
         MaxLength       =   100
         TabIndex        =   51
         Top             =   3945
         Width           =   885
      End
      Begin VB.TextBox TxtMemberMin 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   4995
         MaxLength       =   100
         TabIndex        =   50
         Top             =   3945
         Width           =   885
      End
      Begin VB.TextBox TxtMemberMax 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   6690
         MaxLength       =   100
         TabIndex        =   49
         Top             =   3945
         Width           =   885
      End
      Begin VB.ComboBox CmbPrinters 
         Height          =   315
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   4680
         Width           =   3315
      End
      Begin VB.TextBox TxtX 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   5400
         MaxLength       =   100
         TabIndex        =   45
         Top             =   6030
         Width           =   435
      End
      Begin VB.TextBox TxtY 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   5895
         MaxLength       =   100
         TabIndex        =   44
         Top             =   6030
         Width           =   435
      End
      Begin VB.TextBox TxtConnTimeOut 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5760
         MaxLength       =   2
         TabIndex        =   42
         Top             =   6435
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox TxtSearchDateDifference 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4365
         MaxLength       =   3
         TabIndex        =   37
         Top             =   6030
         Width           =   345
      End
      Begin VB.TextBox TxtHourDifference 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   36
         Top             =   6030
         Width           =   345
      End
      Begin VB.TextBox TxtBlankFooter 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3420
         MaxLength       =   3
         TabIndex        =   35
         Top             =   6390
         Width           =   345
      End
      Begin VB.TextBox TxtStatement 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   270
         MaxLength       =   200
         TabIndex        =   32
         Top             =   7020
         Width           =   10515
      End
      Begin VB.TextBox TxtOrderStatement 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   270
         MaxLength       =   100
         TabIndex        =   31
         Top             =   7590
         Width           =   10515
      End
      Begin SITextBox.Txt TxtStoreID 
         Height          =   315
         Left            =   3285
         TabIndex        =   21
         Tag             =   "NC"
         Top             =   1755
         Width           =   795
         _ExtentX        =   1402
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
         Left            =   4440
         TabIndex        =   22
         Tag             =   "NC"
         Top             =   1755
         Width           =   2415
         _ExtentX        =   4260
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
         Left            =   4080
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1755
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
         MICON           =   "FrmSoftwareDefaultSettings.frx":00C4
         BC              =   12632256
         FC              =   0
      End
      Begin SITextBox.Txt TxtBankMachineID 
         Height          =   315
         Left            =   3285
         TabIndex        =   24
         Top             =   2415
         Width           =   795
         _ExtentX        =   1402
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
      Begin SITextBox.Txt TxtBankMachineName 
         Height          =   315
         Left            =   4440
         TabIndex        =   25
         Top             =   2415
         Width           =   2415
         _ExtentX        =   4260
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
      Begin JeweledBut.JeweledButton BtnBankMachine 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   4080
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2415
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
         MICON           =   "FrmSoftwareDefaultSettings.frx":00E0
         BC              =   12632256
         FC              =   0
      End
      Begin SITextBox.Txt TxtProdDesc1 
         Height          =   315
         Left            =   3255
         TabIndex        =   52
         Top             =   5415
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
         Appearance      =   0
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
         IntegralPoint   =   3
      End
      Begin SITextBox.Txt TxtOrganizationID 
         Height          =   315
         Left            =   3240
         TabIndex        =   64
         Top             =   1110
         Width           =   795
         _ExtentX        =   1402
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
      Begin SITextBox.Txt TxtOrganizationName 
         Height          =   315
         Left            =   4395
         TabIndex        =   65
         Top             =   1110
         Width           =   2415
         _ExtentX        =   4260
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
         Left            =   4035
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1110
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
         MICON           =   "FrmSoftwareDefaultSettings.frx":00FC
         BC              =   12632256
         FC              =   0
      End
      Begin MSComDlg.CommonDialog CD2 
         Left            =   8640
         Top             =   765
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
         Height          =   315
         Left            =   6840
         TabIndex        =   101
         Top             =   6615
         Width           =   1305
         _Version        =   65543
         _ExtentX        =   2302
         _ExtentY        =   556
         _StockProps     =   93
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
         Left            =   8595
         TabIndex        =   102
         Top             =   6615
         Width           =   1305
         _Version        =   65543
         _ExtentX        =   2302
         _ExtentY        =   556
         _StockProps     =   93
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
      Begin SITextBox.Txt TxtSessionID 
         Height          =   315
         Left            =   3285
         TabIndex        =   129
         Top             =   3135
         Width           =   795
         _ExtentX        =   1402
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
      Begin SITextBox.Txt TxtSessionName 
         Height          =   315
         Left            =   4440
         TabIndex        =   130
         Top             =   3135
         Width           =   2415
         _ExtentX        =   4260
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
      Begin JeweledBut.JeweledButton BtnSession 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   4080
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   3135
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
         MICON           =   "FrmSoftwareDefaultSettings.frx":0118
         BC              =   12632256
         FC              =   0
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ohter Charges Name As"
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
         Left            =   10215
         TabIndex        =   226
         Top             =   6165
         Width           =   2040
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Round figure In Search Form"
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
         Left            =   7155
         TabIndex        =   189
         Top             =   8235
         Width           =   2460
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grid Row Height"
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
         Left            =   450
         TabIndex        =   182
         Top             =   3825
         Width           =   1410
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Clssing Fine % on Excess/Short"
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
         Left            =   225
         TabIndex        =   179
         Top             =   8235
         Width           =   3270
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Cost Position"
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
         Left            =   405
         TabIndex        =   170
         Top             =   4410
         Width           =   1650
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top"
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
         Left            =   2610
         TabIndex        =   169
         Top             =   4140
         Width           =   345
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left"
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
         Left            =   2115
         TabIndex        =   168
         Top             =   4140
         Width           =   345
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Late Relax Time"
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
         Left            =   405
         TabIndex        =   165
         Top             =   4905
         Width           =   2265
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode Prefix"
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
         Left            =   405
         TabIndex        =   163
         Top             =   5265
         Width           =   1260
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Charges (%)"
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
         Left            =   405
         TabIndex        =   159
         Top             =   5655
         Width           =   1770
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Printer used For Report"
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
         Left            =   8160
         TabIndex        =   147
         Top             =   5175
         Width           =   2370
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Session Name"
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
         Left            =   4440
         TabIndex        =   133
         Top             =   2925
         Width           =   1215
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Session ID"
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
         Left            =   3285
         TabIndex        =   132
         Top             =   2925
         Width           =   930
      End
      Begin VB.Label LblCaption1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Default Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2430
         TabIndex        =   105
         Top             =   270
         Width           =   2940
      End
      Begin VB.Label LblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
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
         Left            =   8610
         TabIndex        =   104
         Top             =   6390
         Width           =   705
      End
      Begin VB.Label LblFromDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
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
         Left            =   6840
         TabIndex        =   103
         Top             =   6390
         Width           =   885
      End
      Begin VB.Image ImgWaterMark 
         Height          =   1680
         Left            =   8010
         Stretch         =   -1  'True
         Top             =   1845
         Width           =   2565
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Water Mark"
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
         Left            =   7965
         TabIndex        =   69
         Top             =   1620
         Width           =   1005
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
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
         Height          =   195
         Left            =   4995
         TabIndex        =   68
         Top             =   900
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Organization ID"
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
         Left            =   3240
         TabIndex        =   67
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No of Prints in Sale Invoice"
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
         Left            =   2565
         TabIndex        =   57
         Top             =   3690
         Width           =   2355
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Member Min ID"
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
         Left            =   4995
         TabIndex        =   56
         Top             =   3720
         Width           =   1290
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Member Max ID"
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
         Left            =   6690
         TabIndex        =   55
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Printer used in All Reports"
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
         Left            =   3300
         TabIndex        =   54
         Top             =   4410
         Width           =   2235
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Description Show in Report  as"
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
         Left            =   3285
         TabIndex        =   53
         Top             =   5175
         Width           =   3375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Left            =   5535
         TabIndex        =   47
         Top             =   5805
         Width           =   135
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
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
         Left            =   6030
         TabIndex        =   46
         Top             =   5805
         Width           =   135
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connection Timout"
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
         Left            =   4095
         TabIndex        =   43
         Top             =   6480
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Image ImgLogo 
         Height          =   645
         Left            =   9390
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Logo"
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
         Left            =   9000
         TabIndex        =   41
         Top             =   3780
         Width           =   1260
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hour Difference"
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
         Left            =   405
         TabIndex        =   40
         Top             =   6075
         Width           =   1365
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Date Difference"
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
         Left            =   2250
         TabIndex        =   39
         Top             =   6075
         Width           =   2025
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blank Lines in Sale Invoice Footer "
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
         Left            =   405
         TabIndex        =   38
         Top             =   6435
         Width           =   3000
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Footer Statement"
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
         Left            =   270
         TabIndex        =   34
         Top             =   6795
         Width           =   1785
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Footer Statement"
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
         Left            =   270
         TabIndex        =   33
         Top             =   7365
         Width           =   1995
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
         Left            =   3285
         TabIndex        =   30
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Machine ID"
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
         Left            =   3285
         TabIndex        =   29
         Top             =   2205
         Width           =   1485
      End
      Begin VB.Label Label2 
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
         Left            =   4455
         TabIndex        =   28
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DEAB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Machine Name"
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
         TabIndex        =   27
         Top             =   2205
         Width           =   1770
      End
   End
   Begin VB.PictureBox PicSMS 
      BorderStyle     =   0  'None
      Height          =   8100
      Left            =   1875
      ScaleHeight     =   8100
      ScaleWidth      =   12300
      TabIndex        =   77
      Top             =   1875
      Width           =   12300
      Begin VB.TextBox TxtPrefixPhoneNo 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   200
         TabIndex        =   140
         Top             =   1320
         Width           =   1635
      End
      Begin VB.CheckBox ChkAllowSMSOnLogin 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow SMS On Login"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3915
         TabIndex        =   115
         Top             =   5955
         Width           =   3300
      End
      Begin VB.CheckBox ChkAllowSMSThroughDevice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow SMS Through Device"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3915
         TabIndex        =   97
         Top             =   6750
         Width           =   3300
      End
      Begin VB.TextBox TxtWebLinkForSMS 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   270
         MaxLength       =   200
         TabIndex        =   95
         Top             =   4350
         Width           =   10515
      End
      Begin VB.CheckBox ChkAllowSMSOnSave 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow SMS On Save"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3915
         TabIndex        =   94
         Top             =   4770
         Width           =   3300
      End
      Begin VB.CheckBox ChkAllowSMSOnDelete 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow SMS On Delete"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3915
         TabIndex        =   93
         Top             =   5160
         Width           =   3300
      End
      Begin VB.CheckBox ChkAllowSMSOnClear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow SMS On Clear"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3915
         TabIndex        =   92
         Top             =   5565
         Width           =   3300
      End
      Begin VB.CheckBox ChkAllowSMSWithDetail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow SMS  With Detail"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3915
         TabIndex        =   91
         Top             =   6360
         Width           =   3300
      End
      Begin VB.TextBox TxtOwnerMobileNo 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   225
         MaxLength       =   200
         TabIndex        =   80
         Top             =   2115
         Width           =   2235
      End
      Begin JeweledBut.JeweledButton BtnSMSFrm 
         Height          =   420
         Left            =   8415
         TabIndex        =   78
         Top             =   1890
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   741
         TX              =   "SMS"
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
         MICON           =   "FrmSoftwareDefaultSettings.frx":0134
         BC              =   14737632
         FC              =   0
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prefix Phone No like +92"
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
         Left            =   240
         TabIndex        =   141
         Top             =   1080
         Width           =   2130
      End
      Begin MSForms.TextBox TxtCustomerSalesMessage 
         Height          =   1140
         Left            =   240
         TabIndex        =   139
         ToolTipText     =   "Textbox1"
         Top             =   2880
         Width           =   10575
         VariousPropertyBits=   752896027
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "18653;2011"
         SpecialEffect   =   0
         FontName        =   "@Arial Unicode MS"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblCaption1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SMS Default Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   2430
         TabIndex        =   107
         Top             =   270
         Width           =   2910
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web Link For SMS"
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
         Left            =   270
         TabIndex        =   96
         Top             =   4125
         Width           =   1605
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SMS thru API"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4410
         TabIndex        =   83
         Top             =   7425
         Width           =   1860
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SMS thru Device"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4680
         TabIndex        =   82
         Top             =   1620
         Width           =   2340
      End
      Begin VB.Line Line1 
         X1              =   180
         X2              =   10890
         Y1              =   7080
         Y2              =   7080
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner's Mobile No. for Every Sale Bill Message"
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
         Left            =   225
         TabIndex        =   81
         Top             =   1890
         Width           =   4050
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Sales Message"
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
         Left            =   225
         TabIndex        =   79
         Top             =   2655
         Width           =   2130
      End
   End
   Begin VB.PictureBox PicAllow 
      BorderStyle     =   0  'None
      Height          =   8550
      Left            =   -405
      ScaleHeight     =   8550
      ScaleWidth      =   15420
      TabIndex        =   19
      Top             =   2700
      Width           =   15420
      Begin VB.CheckBox ChkAllowDiscountOnSaleDistribution 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Discount On Sale Distribution"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   11070
         TabIndex        =   255
         Top             =   2205
         Width           =   3300
      End
      Begin VB.CheckBox ChkAttendanceNextDayOut 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Attendance Next Day Out"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   11070
         TabIndex        =   254
         Top             =   1800
         Width           =   3345
      End
      Begin VB.CheckBox ChkAllowNegativeStockInBarcodes 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Negative Stock In Barcodes"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   11070
         TabIndex        =   233
         Top             =   1380
         Width           =   3345
      End
      Begin VB.CheckBox ChkDisableQuantityinPOS 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disable Quantity in POS For Standard User"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   11070
         TabIndex        =   232
         Top             =   990
         Width           =   3345
      End
      Begin VB.CheckBox ChkSeperateProductInPOS 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Seperate Product In POS"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   231
         Top             =   975
         Width           =   2220
      End
      Begin VB.CheckBox ChkUseBin 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Bin"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   230
         Top             =   630
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CheckBox ChkTableIDMandatory 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Table ID Mandatory"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4275
         TabIndex        =   204
         Top             =   7170
         Width           =   1740
      End
      Begin VB.CheckBox ChkProductSearchWithStore 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product Search With Store"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   203
         Top             =   6765
         Width           =   2295
      End
      Begin VB.CheckBox ChkEmployeeCommision 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Employee Commision"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   202
         Top             =   1755
         Width           =   1860
      End
      Begin VB.CheckBox ChkChangeQtyPack 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change Qty Pack"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   201
         Top             =   2520
         Width           =   1725
      End
      Begin VB.CheckBox ChkEitherPackORLooseEnter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Either Pack OR Loose Enter"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   200
         Top             =   2880
         Width           =   2445
      End
      Begin VB.CheckBox ChkIsRoundFigure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Round Figure Dist"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   199
         Top             =   3255
         Width           =   1965
      End
      Begin VB.CheckBox ChkIsSingleBarcode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Single Barcode In Products"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   198
         Top             =   3675
         Width           =   2445
      End
      Begin VB.CheckBox ChkDivideRetailWithPacking 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Divide Retail With Packing"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   197
         Top             =   4095
         Width           =   2400
      End
      Begin VB.CheckBox ChkUseMultipleStore 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Multiple Store"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   196
         Top             =   4470
         Width           =   2400
      End
      Begin VB.CheckBox ChkPLSamePR 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "PL Same AS Profit Register"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   195
         Top             =   4860
         Width           =   2295
      End
      Begin VB.CheckBox ChkAdminClosingSaveWhenUserClosingSaved 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Admin Closing Save When User Closing Saved"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   194
         Top             =   7170
         Width           =   3780
      End
      Begin VB.CheckBox ChkLockPurPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lock Purchase Price in Change Price"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   193
         Top             =   6375
         Width           =   3015
      End
      Begin VB.CheckBox ChkUsePurPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Purchase Price"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   192
         Top             =   5625
         Width           =   1815
      End
      Begin VB.CheckBox ChkChangeTransactionDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change Transaction Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   191
         Top             =   6000
         Width           =   2175
      End
      Begin VB.CheckBox ChkAllowBothPackingsareSame 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Both Packings are Same"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   190
         Top             =   1365
         Width           =   2715
      End
      Begin VB.CheckBox ChkAutoPrintinInvoices 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Print in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   186
         Top             =   2145
         Width           =   1815
      End
      Begin VB.CheckBox ChkUsePasswordForm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Password To Open Form For Stanadard User"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   185
         Top             =   7575
         Width           =   3825
      End
      Begin VB.CheckBox chkAllowNegativeOrder 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Negative Order"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   176
         Top             =   2535
         Width           =   1815
      End
      Begin VB.CheckBox ChkCheckStockOnSave 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check Stock On Save"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4275
         TabIndex        =   171
         Top             =   6405
         Width           =   1980
      End
      Begin VB.CheckBox ChkRemarksCompulsory 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remarks Compulsory"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   155
         Top             =   7590
         Width           =   1860
      End
      Begin VB.CheckBox ChkSalePriceLessThanPurchase 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sale Price Less Than Purchase"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   154
         Top             =   2130
         Width           =   2595
      End
      Begin VB.CheckBox ChkSectorCompulsory 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sector Compulsory"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8010
         TabIndex        =   153
         Top             =   7170
         Width           =   1770
      End
      Begin VB.CheckBox ChkAllowDailyBillNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Daily  Bill No"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4275
         TabIndex        =   138
         Top             =   2160
         Width           =   2040
      End
      Begin VB.CheckBox ChkUpdateStockSaleBodyInsert 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update Stock using salebodyinsert Procedure"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   134
         Top             =   7995
         Width           =   3600
      End
      Begin VB.CheckBox ChkEmployeeMandatory 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Employee Mandatory"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4275
         TabIndex        =   125
         Top             =   6780
         Width           =   1845
      End
      Begin VB.CheckBox ChkAutoEnterBeforeQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Enter Before Qty"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   8010
         TabIndex        =   113
         Top             =   5235
         Width           =   2040
      End
      Begin VB.CheckBox ChkSerialCompulsoryinInvoice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Serial Compulsory in Invoice"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4275
         TabIndex        =   112
         Top             =   5265
         Width           =   3300
      End
      Begin VB.CheckBox ChkAutoMoveGridWhenSerialEntered 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Move Grid When Serial Entered"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4275
         TabIndex        =   111
         Top             =   4905
         Width           =   3300
      End
      Begin VB.CheckBox ChkHeaderInfoNotClear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Header Info Not Clear"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4275
         TabIndex        =   109
         Top             =   4500
         Width           =   3300
      End
      Begin VB.CheckBox ChkChangeQtyOnChangedPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change Qty On Changed Price"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4275
         TabIndex        =   99
         Top             =   4080
         Width           =   3300
      End
      Begin VB.CheckBox ChkAllowMonthlyBillNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Monthly  Bill No"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4275
         TabIndex        =   98
         Top             =   1770
         Width           =   2040
      End
      Begin VB.CheckBox ChkAutoEnterQtyintoGridSaleInvoice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Enter Qty into Grid Sale Invoice"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4275
         TabIndex        =   88
         Top             =   3660
         Width           =   3300
      End
      Begin VB.CheckBox ChkAfterRowEditFocusNextGridLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "After Row Edit Focus Next Grid Line"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4275
         TabIndex        =   87
         Top             =   3285
         Width           =   3300
      End
      Begin VB.CheckBox ChkSaveAsNewBill 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Old Bill Converted to New Bill (Save As)"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4275
         TabIndex        =   86
         Top             =   2925
         Width           =   3300
      End
      Begin VB.CheckBox ChkSetEnterKeyGridStockAdjustment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set Enterkey To Grid in Stock Adjustment"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4275
         TabIndex        =   85
         Top             =   2535
         Width           =   3300
      End
      Begin VB.CheckBox ChkAllowContinuousBillNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Continuous Bill No"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4275
         TabIndex        =   84
         Top             =   1380
         Width           =   2040
      End
      Begin VB.CheckBox ChkAllowOrderByCodeinInvoices 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Order By Code in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4275
         TabIndex        =   75
         Top             =   990
         Width           =   2580
      End
      Begin VB.CheckBox ChkAutoApplyPartyLastDiscount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Apply Party Last Discount in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4275
         TabIndex        =   74
         Top             =   6030
         Width           =   3390
      End
      Begin VB.CheckBox ChkAlertAllocateProduct 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alert Allocate Product"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   1755
         Width           =   1860
      End
      Begin VB.CheckBox ChkSeperateProductWithPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Seperate Product With Price"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   3690
         Width           =   2355
      End
      Begin VB.CheckBox ChkDisableAutoPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disable Auto Print"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   990
         Width           =   1590
      End
      Begin VB.CheckBox ChkAutoPrintSaleOrder 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Print Sale Order"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4275
         TabIndex        =   18
         Top             =   7995
         Width           =   1815
      End
      Begin VB.CheckBox ChkPrintKitchenInoices 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print Kitchen Inoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4275
         TabIndex        =   17
         Top             =   7575
         Width           =   1770
      End
      Begin VB.CheckBox ChkSaleInProduction 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Sales In Production"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   3300
         Width           =   2130
      End
      Begin VB.CheckBox ChkCashReceived 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alllow Auto Cash Received in Sale Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   6030
         Width           =   3345
      End
      Begin VB.CheckBox ChkNegativeSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Negative Sales"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   1380
         Width           =   1815
      End
      Begin VB.CheckBox ChkAutoApplyPartyLastPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Apply Party Last Price in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4275
         TabIndex        =   16
         Top             =   5640
         Width           =   3030
      End
      Begin VB.CheckBox ChkCostVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable Cost Function Keys in Sale Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   5640
         Width           =   3300
      End
      Begin VB.CheckBox ChkLaserPrintofSaleInvoice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Laser Print of Sale Invoice Half"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   4470
         Width           =   2535
      End
      Begin VB.CheckBox ChkPrintHeadersSaleInvoice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print Headers in Sale Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   4080
         Width           =   2400
      End
      Begin VB.CheckBox ChkSystemDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Set System Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   2925
         Width           =   1950
      End
      Begin VB.CheckBox ChkProperCase 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Product Name in Proper Case"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   4860
         Width           =   2805
      End
      Begin VB.CheckBox ChkProductSearchOpenInPreviousState 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Product Search Open in Previous State"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   6405
         Width           =   3525
      End
      Begin VB.CheckBox ChkDiscountAllowed 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Discount Allowed For Standard User"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   5250
         Width           =   2895
      End
      Begin VB.CheckBox ChkChangePrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Administrator can change Price in Sale Transections"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   6795
         Width           =   4020
      End
      Begin VB.Label LblCaption1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Default Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   201
         Left            =   2430
         TabIndex        =   63
         Top             =   270
         Width           =   3030
      End
   End
   Begin VB.PictureBox PicEmail 
      BorderStyle     =   0  'None
      Height          =   6450
      Left            =   1980
      ScaleHeight     =   6450
      ScaleWidth      =   12300
      TabIndex        =   235
      Top             =   1980
      Width           =   12300
      Begin VB.TextBox TxtActivityActionNo 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1980
         MaxLength       =   100
         TabIndex        =   250
         Top             =   5085
         Width           =   3990
      End
      Begin VB.CheckBox ChkExportReportASPDF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Export Report AS PDF"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3780
         TabIndex        =   248
         Top             =   1305
         Width           =   2580
      End
      Begin VB.CheckBox ChkUseEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Email"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1995
         TabIndex        =   247
         Top             =   1305
         Width           =   1590
      End
      Begin VB.TextBox txtEmailPwd 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   6120
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   245
         Tag             =   "Admin"
         Top             =   1965
         Width           =   2325
      End
      Begin VB.TextBox TxtFromEmail 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1995
         MaxLength       =   100
         TabIndex        =   244
         Top             =   1965
         Width           =   3990
      End
      Begin VB.ComboBox CmbSMTPServerAddress 
         Height          =   315
         Left            =   1995
         Style           =   2  'Dropdown List
         TabIndex        =   238
         Top             =   3510
         Width           =   3990
      End
      Begin VB.TextBox TxtToEmail 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1995
         MaxLength       =   100
         TabIndex        =   237
         Top             =   2640
         Width           =   3990
      End
      Begin SITextBox.Txt TxtPortNo 
         Height          =   315
         Left            =   6075
         TabIndex        =   243
         Top             =   3510
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         Appearance      =   0
         MaxLength       =   9
         Text            =   "25"
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
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Action No. For User Activity Report"
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
         Left            =   1980
         TabIndex        =   251
         Top             =   4860
         Width           =   3015
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note:- For Gmail account You should Manage your Google Account  Security ""Less secure app access"" On"
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
         Left            =   1980
         TabIndex        =   249
         Top             =   4320
         Width           =   9135
      End
      Begin VB.Label LblCaption1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Index           =   5
         Left            =   6165
         TabIndex        =   246
         Top             =   1710
         Width           =   825
      End
      Begin VB.Label LblCaption1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port No"
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
         Index           =   8
         Left            =   6075
         TabIndex        =   242
         Top             =   3285
         Width           =   660
      End
      Begin VB.Label LblCaption1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP server address"
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
         Index           =   7
         Left            =   1995
         TabIndex        =   241
         Top             =   3240
         Width           =   1830
      End
      Begin VB.Label LblCaption1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
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
         Index           =   6
         Left            =   1995
         TabIndex        =   240
         Top             =   2415
         Width           =   270
      End
      Begin VB.Label LblCaption1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Index           =   4
         Left            =   1995
         TabIndex        =   239
         Top             =   1710
         Width           =   420
      End
      Begin VB.Label LblCaption1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email Default Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2430
         TabIndex        =   236
         Top             =   270
         Width           =   3045
      End
   End
   Begin VB.PictureBox PicShowHide 
      BorderStyle     =   0  'None
      Height          =   8055
      Left            =   45
      ScaleHeight     =   8055
      ScaleWidth      =   15405
      TabIndex        =   1
      Top             =   1395
      Width           =   15405
      Begin VB.CheckBox ChkShowMultiBranches 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Multi Branches"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   10890
         TabIndex        =   256
         Top             =   5130
         Width           =   3525
      End
      Begin VB.CheckBox ChkShowBarcodeDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Barcode Description"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10890
         TabIndex        =   253
         Top             =   4365
         Width           =   2445
      End
      Begin VB.CheckBox ChkShowDiscPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Only Disc Price in Price Checker"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   10890
         TabIndex        =   252
         Top             =   4725
         Width           =   3525
      End
      Begin VB.CheckBox ChkShowBankInTransection 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Bank in Transections For Standard User"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10890
         TabIndex        =   229
         Top             =   3975
         Width           =   3615
      End
      Begin VB.CheckBox ChkisShowPublisher 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Publisher in Product"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   228
         Top             =   3240
         Width           =   2445
      End
      Begin VB.CheckBox ChkShowChangeRetailinPurchaseInvoice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Change Retail in Purchase Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10890
         TabIndex        =   227
         Top             =   3225
         Width           =   3285
      End
      Begin VB.CheckBox ChkStoreVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Stores in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   224
         Top             =   1365
         Width           =   2085
      End
      Begin VB.CheckBox ChkTableVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Tables in Sale Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   223
         Top             =   2865
         Width           =   2400
      End
      Begin VB.CheckBox ChkEmployeeVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Employee in Sale Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   222
         Top             =   5865
         Width           =   2625
      End
      Begin VB.CheckBox ChkMemberVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Member in Sale Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   221
         Top             =   4365
         Width           =   2490
      End
      Begin VB.CheckBox ChkHideSaleAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hide Amount in Previous Sale Transections For Standard User"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   220
         Top             =   8115
         Width           =   4740
      End
      Begin VB.CheckBox ChkHidePurchaseAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hide Amount in Purchase Transections For Standard User"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   219
         Top             =   7740
         Width           =   4470
      End
      Begin VB.CheckBox ChkTag 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Tag in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   218
         Top             =   990
         Width           =   1860
      End
      Begin VB.CheckBox ChkPreviousBalanceVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Previous Balance in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   217
         Top             =   6240
         Width           =   2850
      End
      Begin VB.CheckBox ChkManualBillNoVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Manual Bill in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   216
         Top             =   2115
         Width           =   2355
      End
      Begin VB.CheckBox ChkRemarksVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Remarks in Sale Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   215
         Top             =   3990
         Width           =   2490
      End
      Begin VB.CheckBox ChkOrganizationVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Organization in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   214
         Top             =   3240
         Width           =   2445
      End
      Begin VB.CheckBox ChkFright 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Freight Option in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   213
         Top             =   4740
         Width           =   2580
      End
      Begin VB.CheckBox ChkShowCodeInHalfPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Code In Half Print Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   212
         Top             =   5115
         Width           =   2580
      End
      Begin VB.CheckBox ChkShowSerialInHalfPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Serial In Half Print Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   211
         Top             =   5490
         Width           =   2580
      End
      Begin VB.CheckBox ChkInvType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Types in Sales Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   210
         Top             =   2490
         Width           =   2445
      End
      Begin VB.CheckBox ChkBatchNoVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show BatchNo in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   209
         Top             =   1740
         Width           =   2310
      End
      Begin VB.CheckBox ChkSaleOrderVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Sale Order in Sale Invoice Search"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   208
         Top             =   6990
         Width           =   3165
      End
      Begin VB.CheckBox ChkShowRetailinPurchaseReturnPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Retail in Purchase Return Print"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   207
         Top             =   6615
         Width           =   2985
      End
      Begin VB.CheckBox ChkShowRawMaterialProductInSaleInvoices 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Raw Material Product In Sale Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   206
         Top             =   7365
         Width           =   3525
      End
      Begin VB.CheckBox ChkOrganizationMandatory 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Organization Mandatory"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   205
         Top             =   3615
         Width           =   2445
      End
      Begin VB.CheckBox ChkShowPurchaseProfit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Purchase Profit"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   187
         Top             =   4740
         Width           =   2160
      End
      Begin VB.CheckBox ChkSearchCodeInGrid 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Search Code In Grid"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   184
         Top             =   7365
         Width           =   2430
      End
      Begin VB.CheckBox ChkShowAllStoreStock 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show All Store Stock"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10890
         TabIndex        =   183
         Top             =   2850
         Width           =   1890
      End
      Begin VB.CheckBox ChkShowPurPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Purchase Price"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   180
         Top             =   3240
         Width           =   2025
      End
      Begin VB.CheckBox ChkShowAddBarCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Add BarCode"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10890
         TabIndex        =   177
         Top             =   2475
         Width           =   2025
      End
      Begin VB.CheckBox ChkShowExpiryInvoice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Expiry Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10890
         TabIndex        =   175
         Top             =   3600
         Width           =   1905
      End
      Begin VB.CheckBox ChkShowSaleTax 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Sale Tax"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10890
         TabIndex        =   174
         Top             =   2082
         Width           =   1485
      End
      Begin VB.CheckBox chkShowStockPriceChecker 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Stock Price Checker"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   173
         Top             =   7740
         Width           =   2430
      End
      Begin VB.CheckBox ChkShowDiscPurPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Disc PurPrice in Chage Price"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   172
         Top             =   6975
         Width           =   3105
      End
      Begin VB.CheckBox ChkShowReSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show ReSale"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   161
         Top             =   6600
         Width           =   2025
      End
      Begin VB.CheckBox ChkShowTimeFilterinReport 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Time Filter in Report"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   160
         Top             =   3600
         Width           =   2430
      End
      Begin VB.CheckBox ChkShowAllPrices 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show All Prices in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   157
         Top             =   6990
         Width           =   2295
      End
      Begin VB.CheckBox ChkShowDispatchDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Terms & Dispatch Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   156
         Top             =   6225
         Width           =   2475
      End
      Begin VB.CheckBox ChkShowGrandTotalinSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Grand Total in Search"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   152
         Top             =   5850
         Width           =   2655
      End
      Begin VB.CheckBox ChkShowChangePriceOnSavePI 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Change Price On Save PI"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   151
         Top             =   5475
         Width           =   2655
      End
      Begin VB.CheckBox chkisShowListPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show List Price"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   150
         Top             =   5100
         Width           =   1485
      End
      Begin VB.CheckBox ChkShowHistoryofAllCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show History of All Customer"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   149
         Top             =   4725
         Width           =   2445
      End
      Begin VB.CheckBox ChkShowBarCodeQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show BarCode Qty"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   148
         Top             =   4350
         Width           =   1725
      End
      Begin VB.CheckBox ChkShowBatchPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Batch Print"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   145
         Top             =   3975
         Width           =   1725
      End
      Begin VB.CheckBox ChkShowSC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show SC"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10890
         TabIndex        =   144
         Top             =   975
         Width           =   1365
      End
      Begin VB.CheckBox ChkShowOffer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Offer"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10890
         TabIndex        =   143
         Top             =   1725
         Width           =   1245
      End
      Begin VB.CheckBox ChkShowBonus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Bonus"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10890
         TabIndex        =   142
         Top             =   1350
         Width           =   1245
      End
      Begin VB.CheckBox ChkShowTradeOffer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Trade Offer"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   135
         Top             =   6615
         Width           =   2805
      End
      Begin VB.CheckBox ChkShowSession 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Session"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   128
         Top             =   6240
         Width           =   2805
      End
      Begin VB.CheckBox ChkShowSavedStock 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Saved Stock in Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   127
         Top             =   5115
         Width           =   2805
      End
      Begin VB.CheckBox ChkIsGrossQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Gross Qty in Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   126
         Top             =   5865
         Width           =   2805
      End
      Begin VB.CheckBox ChkShowColourSize 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Colour and Size in Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   124
         Top             =   5490
         Width           =   2805
      End
      Begin VB.CheckBox ChkisShowVendor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Vendor in Product"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   123
         Top             =   2850
         Width           =   2445
      End
      Begin VB.CheckBox ChkisShowItemDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "how Item Desc in Product"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   120
         Top             =   1725
         Width           =   2445
      End
      Begin VB.CheckBox ChkisShowOther 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Other in Product"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   122
         Top             =   2475
         Width           =   2445
      End
      Begin VB.CheckBox chkisShowSeason 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Season in Product"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   121
         Top             =   2100
         Width           =   2445
      End
      Begin VB.CheckBox ChkisShowSubDepartment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Sub Department in Product"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   119
         Top             =   1350
         Width           =   2895
      End
      Begin VB.CheckBox ChkisShowDepartment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Department in Product"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   118
         Top             =   975
         Width           =   2445
      End
      Begin VB.CheckBox ChkShowStockFromTableGridDataMovement 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Stock From Table Grid Data Movement"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   117
         Top             =   7350
         Width           =   3705
      End
      Begin VB.CheckBox ChkShowBarcodeProductSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Barcode Product Search"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   116
         Top             =   4365
         Width           =   2715
      End
      Begin VB.CheckBox ChkShowProdProfit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show PPxxxxx"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   110
         Top             =   3990
         Width           =   1410
      End
      Begin VB.CheckBox ChkShowSyllabus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Syllabus"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   108
         Top             =   3615
         Width           =   1545
      End
      Begin VB.CheckBox ChkShowPromiseDateInSalaPurchase 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Promise Date In Sale Purchase"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   90
         Top             =   2865
         Width           =   3075
      End
      Begin VB.CheckBox ChkShowWholeSaleMargin 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Whole Sale Margin"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   89
         Top             =   2490
         Width           =   2265
      End
      Begin VB.CheckBox ChkHideClearButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hide Clear Botton"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4005
         TabIndex        =   73
         Top             =   2115
         Width           =   1620
      End
      Begin VB.CheckBox ChkShowWarrantyinSaleInvoice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Warranty in Sale Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   72
         Top             =   1740
         Width           =   2535
      End
      Begin VB.CheckBox ChkShowLastInvoiceMsgAtSave 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Last Invoice Msg At Save"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   71
         Top             =   1365
         Width           =   2670
      End
      Begin VB.CheckBox ChkQuantityinBarcodes 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Quantity in Barcodes"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4005
         TabIndex        =   70
         Top             =   990
         Width           =   2265
      End
      Begin VB.Label LblCaption1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show / Hide Default Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2430
         TabIndex        =   106
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Software Default Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2700
      TabIndex        =   0
      Top             =   270
      Width           =   3465
   End
   Begin VB.Image ImgExit 
      Height          =   360
      Left            =   13140
      Top             =   998
      Width           =   330
   End
End
Attribute VB_Name = "FrmSoftwareDefaultSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vFlag As Boolean
Dim sSql As String

Private Function FunSelectBankMachine(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchBankMachine.Show vbModal, Me
        If SchBankMachine.ParaOutBankMachineID = "" Then FunSelectBankMachine = False: Exit Function
        TxtBankMachineID.Text = SchBankMachine.ParaOutBankMachineID
    End If
    '---------------------------
    vStrSQL = " Select * FROM BankMachines where BankMachineID=" & Val(TxtBankMachineID.Text)
    CN.CursorLocation = adUseClient
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtBankMachineName.Text = !BankMachineName
          FunSelectBankMachine = True
          .Close
          Exit Function
      Else
          FunSelectBankMachine = False
          .Close
          TxtBankMachineID.Text = ""
          TxtBankMachineName.Text = ""
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
    CN.CursorLocation = adUseClient
    With CN.Execute(vStrSQL)
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

Private Sub BtnAllow_Click()
   On Error GoTo ErrorHandler
'   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   PicAllow.Visible = True
   PicDefault.Visible = False
   PicSMS.Visible = False
   PicShowHide.Visible = False
   PicEmail.Visible = False
   
   BtnAllow.FontBold = True
   BtnSMS.FontBold = False
   BtnDefault.FontBold = False
   BtnShowHide.FontBold = False
   
   PicAllow.Left = 0
   PicAllow.Top = 100
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnDefault_Click()
   On Error GoTo ErrorHandler
'   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   PicDefault.Visible = True
   PicSMS.Visible = False
   PicAllow.Visible = False
   PicShowHide.Visible = False
   PicEmail.Visible = False
   
   
   BtnDefault.FontBold = True
   BtnSMS.FontBold = False
   BtnAllow.FontBold = False
   BtnShowHide.FontBold = False
   
   PicDefault.Left = 125
   PicDefault.Top = 100
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnEmail_Click()
 On Error GoTo ErrorHandler
   PicEmail.Visible = True
   PicSMS.Visible = False
   PicDefault.Visible = False
   PicAllow.Visible = False
   PicShowHide.Visible = False
   
   BtnEmail.FontBold = True
   BtnSMS.FontBold = False
   BtnDefault.FontBold = False
   BtnAllow.FontBold = False
   BtnShowHide.FontBold = False
   
   PicEmail.Left = 125
   PicEmail.Top = 100
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSession_Click()
   If FunSelectSession(ssButton, False) = True Then
      TxtNoofPrints.SetFocus
   Else
      TxtSessionID.SetFocus
   End If
End Sub

Private Sub BtnShowHide_Click()
   On Error GoTo ErrorHandler
'   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   
   PicShowHide.Visible = True
   PicSMS.Visible = False
   PicDefault.Visible = False
   PicAllow.Visible = False
   PicEmail.Visible = False
   
   BtnShowHide.FontBold = True
   BtnSMS.FontBold = False
   BtnDefault.FontBold = False
   BtnAllow.FontBold = False
   
   PicShowHide.Left = 0
   PicShowHide.Top = 100
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSMS_Click()
   On Error GoTo ErrorHandler
'   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   PicSMS.Visible = True
   PicDefault.Visible = False
   PicAllow.Visible = False
   PicShowHide.Visible = False
   PicEmail.Visible = False
   
   BtnSMS.FontBold = True
   BtnDefault.FontBold = False
   BtnAllow.FontBold = False
   BtnShowHide.FontBold = False
   
   PicSMS.Left = 125
   PicSMS.Top = 100
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSMSFrm_Click()
   FrmSMS.Show
End Sub

Private Sub ChkStoreVisible_Click()
   On Error GoTo ErrorHandler
   TxtStoreID.Enabled = ChkStoreVisible.Value = 1
   TxtStoreName.Enabled = TxtStoreID.Enabled
   BtnStore.Enabled = TxtStoreID.Enabled
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   On Error GoTo ErrorHandler
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbSMTPServerAddress_Click()
On Error GoTo ErrorHandler
   If CmbSMTPServerAddress.Text = "smtp.gmail.com" Then
      TxtPortNo.Text = 465
   ElseIf CmbSMTPServerAddress.Text = "smtp.live.com" Then
      TxtPortNo.Text = 25
   Else
      TxtPortNo.Text = 25
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then TxtStoreID.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtBankMachineID.SetFocus
         Case TxtBankMachineID.Name: If FunSelectBankMachine(ssFunctionKey, False) = True Then TxtNoofPrints.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   If FunValidation = False Then Exit Sub
   Dim vPrinter() As String
   Dim vPrinter2() As String
   vPrinter = Split(CmbPrinters.Text, ",")
   vPrinter2 = Split(Cmb2ndPrinters.Text, ",")
   CN.CursorLocation = adUseClient
   Call SaveLogo
   Call SaveWaterMark
   
   ''''' Form Default Settings '''''''''''
   vPrinter = Split(CmbPrinters.Text, ",")
   sSql = "select * from FormDefaultSetting Where FormType = 'Software Default Setting' and LocalComputerName = '" & LocalComputerName & "'"
   If CN.Execute(sSql).EOF Then
      sSql = "Insert into FormDefaultSetting (LocalComputerName, FormType, Size, DeviceName, DriverName, Port, IsPreview ) Values ('" & LocalComputerName & "', 'Software Default Setting','" & CmbPrinters.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "'," & 0 & ")"
   Else
      sSql = "Update FormDefaultSetting set Size = '" & CmbPrinters.Text & "', DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "', IsPreview = " & 0 & " Where FormType = 'Software Default Setting' and LocalComputerName = '" & LocalComputerName & "'"
   End If
   CN.Execute sSql
   ''''''''''''''''''''''''''''''''''''''''''''
   'Allow
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkNegativeSale.Value & "' where RegistryKey = 'NegativeSale'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAllowOrderByCodeinInvoices.Value & "' where RegistryKey = 'AllowOrderByCodeinInvoices'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAllowMonthlyBillNo.Value & "' where RegistryKey = 'AllowMonthlyBillNo'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAllowContinuousBillNo.Value & "' where RegistryKey = 'AllowContinuousBillNo'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAllowDailyBillNo.Value & "' where RegistryKey = 'AllowDailyBillNo'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAllowBothPackingsareSame.Value & "' where RegistryKey = 'AllowBothPackingsareSame'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAlertAllocateProduct.Value & "' where RegistryKey = 'AlertAllocateProduct'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkSystemDate.Value & "' where RegistryKey = 'SystemDate'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkSaleInProduction.Value & "' where RegistryKey = 'SaleInProduction'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkPrintHeadersSaleInvoice.Value & "' where RegistryKey = 'PrintHeadersSaleInvoice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkLaserPrintofSaleInvoice.Value & "' where RegistryKey = 'LaserPrintofSaleInvoice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkProperCase.Value & "' where RegistryKey = 'ProperCase'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkDiscountAllowed.Value & "' where RegistryKey = 'DiscAllowed'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkCostVisible.Value & "' where RegistryKey = 'CostVisible'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkCashReceived.Value & "' where RegistryKey = 'CashReceived'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkProductSearchOpenInPreviousState.Value & "' where RegistryKey = 'ProductSearchOpenInPreviousState'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkChangePrice.Value & "' where RegistryKey = 'ChangePrice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAutoApplyPartyLastPrice.Value & "' where RegistryKey = 'AutoApplyPartyLastPrice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAutoApplyPartyLastDiscount.Value & "' where RegistryKey = 'AutoApplyPartyLastDiscount'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkPrintKitchenInoices.Value & "' where RegistryKey = 'PrintKitchenInoices'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAutoPrintSaleOrder.Value & "' where RegistryKey = 'AutoPrintSaleOrder'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkDisableAutoPrint.Value & "' where RegistryKey = 'HideAutoPrint'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkDisableQuantityinPOS.Value & "' where RegistryKey = 'DisableQuantityinPOS'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAllowNegativeStockInBarcodes.Value & "' where RegistryKey = 'AllowNegativeStockInBarcodes'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAllowDiscountOnSaleDistribution.Value & "' where RegistryKey = 'AllowDiscountOnSaleDistribution'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkSeperateProductWithPrice.Value & "' where RegistryKey = 'SeperateProductWithPrice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkSeperateProductInPOS.Value & "' where RegistryKey = 'SeperateProductInPOS'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkSetEnterKeyGridStockAdjustment.Value & "' where RegistryKey = 'SetEnterKeyGridStockAdjustment'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkSaveAsNewBill.Value & "' where RegistryKey = 'SaveAsNewBill'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAfterRowEditFocusNextGridLine.Value & "' where RegistryKey = 'AfterRowEditFocusNextGridLine'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkHideClearButton.Value & "' where RegistryKey = 'HideClearButton'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAutoEnterQtyintoGridSaleInvoice.Value & "' where RegistryKey = 'AutoEnterQtyintoGridSaleInvoice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAutoEnterBeforeQty.Value & "' where RegistryKey = 'AutoEnterBeforeQty'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAllowSMSOnSave.Value & "' where RegistryKey = 'AllowSMSOnSave'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAllowSMSOnDelete.Value & "' where RegistryKey = 'AllowSMSOnDelete'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAllowSMSOnClear.Value & "' where RegistryKey = 'AllowSMSOnClear'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAllowSMSOnLogin.Value & "' where RegistryKey = 'AllowSMSOnLogin'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAllowSMSWithDetail.Value & "' where RegistryKey = 'AllowSMSWithDetail'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAllowSMSThroughDevice.Value & "' where RegistryKey = 'AllowSMSThroughDevice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkChangeQtyOnChangedPrice.Value & "' where RegistryKey = 'ChangeQtyOnChangedPrice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkHeaderInfoNotClear.Value & "' where RegistryKey = 'HeaderInfoNotClear'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAutoMoveGridWhenSerialEntered.Value & "' where RegistryKey = 'AutoMoveGridWhenSerialEntered'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkSerialCompulsoryinInvoice.Value & "' where RegistryKey = 'SerialCompulsoryinInvoice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkEmployeeMandatory.Value & "' where RegistryKey = 'EmployeeMandatory'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkTableIDMandatory.Value & "' where RegistryKey = 'TableIDMandatory'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkUpdateStockSaleBodyInsert.Value & "' where RegistryKey = 'UpdateStockSaleBodyInsert'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkEmployeeCommision.Value & "' where RegistryKey = 'EmployeeCommision'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkSalePriceLessThanPurchase.Value & "' where RegistryKey = 'SalePriceLessThanPurchase'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkChangeQtyPack.Value & "' where RegistryKey = 'ChangeQtyPack'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkEitherPackORLooseEnter.Value & "' where RegistryKey = 'EitherPackORLooseEnter'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkIsSingleBarcode.Value & "' where RegistryKey = 'IsSingleBarcode'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkSectorCompulsory.Value & "' where RegistryKey = 'SectorCompulsory'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkRemarksCompulsory.Value & "' where RegistryKey = 'RemarksCompulsory'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkDivideRetailWithPacking.Value & "' where RegistryKey = 'DivideRetailWithPacking'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkUseMultipleStore.Value & "' where RegistryKey = 'UseMultipleStore'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkPLSamePR.Value & "' where RegistryKey = 'PLSamePR'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAdminClosingSaveWhenUserClosingSaved.Value & "' where RegistryKey = 'AdminClosingSaveWhenUserClosingSaved'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkCheckStockOnSave.Value & "' where RegistryKey = 'CheckStockOnSave'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAutoPrintinInvoices.Value & "' where RegistryKey = 'AutoPrintinInvoices'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkUsePurPrice.Value & "' where RegistryKey = 'UsePurPrice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkChangeTransactionDate.Value & "' where RegistryKey = 'ChangeTransactionDate'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkUseBin.Value & "' where RegistryKey = 'UseBin'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkAttendanceNextDayOut.Value & "' where RegistryKey = 'AttendanceNextDayOut'")
   
   'Show Hide
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkBatchNoVisible.Value & "' where RegistryKey = 'BatchNoVisible'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkQuantityinBarcodes.Value & "' where RegistryKey = 'QuantityinBarcodes'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowLastInvoiceMsgAtSave.Value & "' where RegistryKey = 'ShowLastInvoiceMsgAtSave'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkTag.Value & "' where RegistryKey = 'Tag'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkStoreVisible.Value & "' where RegistryKey = 'StoreVisible'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkManualBillNoVisible.Value & "' where RegistryKey = 'ManualBillNoVisible'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkTableVisible.Value & "' where RegistryKey = 'TableVisible'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkOrganizationVisible.Value & "' where RegistryKey = 'OrganizationVisible'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkRemarksVisible.Value & "' where RegistryKey = 'RemarksVisible'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkMemberVisible.Value & "' where RegistryKey = 'MemberVisible'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkSaleOrderVisible.Value & "' where RegistryKey = 'SaleOrderVisible'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkFright.Value & "' where RegistryKey = 'FreightVisible'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkEmployeeVisible.Value & "' where RegistryKey = 'EmpVisible'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkPreviousBalanceVisible.Value & "' where RegistryKey = 'PreviousBalanceVisible'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkHideSaleAmount.Value & "' where RegistryKey = 'HideSaleAmount'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkHidePurchaseAmount.Value & "' where RegistryKey = 'HidePurchaseAmount'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowBankInTransection.Value & "' where RegistryKey = 'ShowBankInTransection'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowWarrantyinSaleInvoice.Value & "' where RegistryKey = 'ShowWarrantyinSaleInvoice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowCodeInHalfPrint.Value & "' where RegistryKey = 'ShowCode'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowSerialInHalfPrint.Value & "' where RegistryKey = 'ShowSerial'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowRetailinPurchaseReturnPrint.Value & "' where RegistryKey = 'ShowRetailinPurchaseReturnPrint'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowRawMaterialProductInSaleInvoices.Value & "' where RegistryKey = 'ShowRawMaterialProductInSaleInvoices'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkInvType.Value & "' where RegistryKey = 'InvType'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowWholeSaleMargin.Value & "' where RegistryKey = 'ShowWholeSaleMargin'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowPromiseDateInSalaPurchase.Value & "' where RegistryKey = 'ShowPromiseDateInSalaPurchase'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowSyllabus.Value & "' where RegistryKey = 'ShowSyllabus'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowProdProfit.Value & "' where RegistryKey = 'ShowProdProfit'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowStockFromTableGridDataMovement.Value & "' where RegistryKey = 'ShowStockFromTableGridDataMovement'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowBarcodeProductSearch.Value & "' where RegistryKey = 'ShowBarcodeProductSearch'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkOrganizationMandatory.Value & "' where RegistryKey = 'OrganizationMandatory'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkisShowPublisher.Value & "' where RegistryKey = 'isShowPublisher'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & chkisShowListPrice.Value & "' where RegistryKey = 'isShowListPrice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkisShowDepartment.Value & "' where RegistryKey = 'isShowDepartment'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkisShowSubDepartment.Value & "' where RegistryKey = 'isShowSubDepartment'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & chkisShowSeason.Value & "' where RegistryKey = 'isShowSeason'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkisShowItemDesc.Value & "' where RegistryKey = 'isShowItemDesc'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkisShowOther.Value & "' where RegistryKey = 'isShowOther'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkisShowVendor.Value & "' where RegistryKey = 'isShowVendor'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowColourSize.Value & "' where RegistryKey = 'ShowColourSize'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkIsGrossQty.Value & "' where RegistryKey = 'IsGrossQty'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowSavedStock.Value & "' where RegistryKey = 'ShowSavedStock'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowAllPrices.Value & "' where RegistryKey = 'ShowAllPrices'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowChangeRetailinPurchaseInvoice.Value & "' where RegistryKey = 'ShowChangeRetailinPurchaseInvoice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowSession.Value & "' where RegistryKey = 'ShowSession'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowTradeOffer.Value & "' where RegistryKey = 'ShowTradeOffer'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowExpiryInvoice.Value & "' where RegistryKey = 'ShowExpiryInvoice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowBonus.Value & "' where RegistryKey = 'ShowBonus'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowOffer.Value & "' where RegistryKey = 'ShowOffer'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowSaleTax.Value & "' where RegistryKey = 'ShowSaleTax'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowSC.Value & "' where RegistryKey = 'ShowSC'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtChargesName.Text & "' where RegistryKey = 'ChargesName'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowBatchPrint.Value & "' where RegistryKey = 'ShowBatchPrint'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowBarCodeQty.Value & "' where RegistryKey = 'ShowBarCodeQty'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowHistoryofAllCustomer.Value & "' where RegistryKey = 'ShowHistoryofAllCustomer'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkIsRoundFigure.Value & "' where RegistryKey = 'IsRoundFigure'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowChangePriceOnSavePI.Value & "' where RegistryKey = 'ShowChangePriceOnSavePI'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowGrandTotalinSearch.Value & "' where RegistryKey = 'ShowGrandTotalinSearch'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowDispatchDate.Value & "' where RegistryKey = 'ShowDispatchDate'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowTimeFilterinReport.Value & "' where RegistryKey = 'TimeWiseReport'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowReSale.Value & "' where RegistryKey = 'ShowReSale'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkLockPurPrice.Value & "' where RegistryKey = 'LockPurPrice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowDiscPurPrice.Value & "' where RegistryKey = 'ShowDiscPurPrice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & chkShowStockPriceChecker.Value & "' where RegistryKey = 'ShowStockPriceChecker'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & chkAllowNegativeOrder.Value & "' where RegistryKey = 'AllowNegativeOrder'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowAddBarCode.Value & "' where RegistryKey = 'ShowAddBarCode'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowPurPrice.Value & "' where RegistryKey = 'ShowPurPrice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkProductSearchWithStore.Value & "' where RegistryKey = 'ProductSearchWithStore'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowAllStoreStock.Value & "' where RegistryKey = 'ShowAllStoreStock'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkSearchCodeInGrid.Value & "' where RegistryKey = 'SearchCodeInGrid'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkUsePasswordForm.Value & "' where RegistryKey = 'UsePasswordForm'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowPurchaseProfit.Value & "' where RegistryKey = 'ShowPurchaseProfit'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowBarcodeDesc.Value & "' where RegistryKey = 'ShowBarcodeDesc'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowDiscPrice.Value & "' where RegistryKey = 'ShowDiscPrice'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkShowMultiBranches.Value & "' where RegistryKey = 'ShowMultiBranches'")
   
   'Default Values
   CN.Execute ("UPDATE sysindexs Set Value = '" & IIf(Trim(TxtOrganizationID.Text) = "", Null, Val(TxtOrganizationID.Text)) & "' where RegistryKey = 'OrganizationID'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & IIf(Trim(TxtStoreID.Text) = "", Null, Val(TxtStoreID.Text)) & "' where RegistryKey = 'StoreID'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & IIf(Trim(TxtBankMachineID.Text) = "", Null, Val(TxtBankMachineID.Text)) & "' where RegistryKey = 'BankMachineID'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & IIf(Trim(TxtSessionID.Text) = "", Null, Val(TxtSessionID.Text)) & "' where RegistryKey = 'SessionID'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtNoofPrints.Text) & "' where RegistryKey = 'NoofPrints'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtMemberMin.Text) & "' where RegistryKey = 'MemberMin'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtMemberMax.Text) & "' where RegistryKey = 'MemberMax'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter(0) & "' where RegistryKey = 'DeviceName'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter(1) & "' where RegistryKey = 'DriverName'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter(2) & "' where RegistryKey = 'Port'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter2(0) & "' where RegistryKey = 'DeviceName2'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter2(1) & "' where RegistryKey = 'DriverName2'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter2(2) & "' where RegistryKey = 'Port2'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtProdDesc1.Text & "' where RegistryKey = 'ProdDesc1'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtPackingChargesPer.Text) & "' where RegistryKey = 'PackingChargesPer'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtHourDifference.Text) & "' where RegistryKey = 'HourDifference'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtBarcodePrefix.Text) & "' where RegistryKey = 'BarCodePrefix'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtSearchDateDifference.Text) & "' where RegistryKey = 'SearchDateDifference'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtX.Text) & "' where RegistryKey = 'X'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtY.Text) & "' where RegistryKey = 'Y'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtCostX.Text) & "' where RegistryKey = 'CostX'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtCostY.Text) & "' where RegistryKey = 'CostY'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtBlankFooter.Text) & "' where RegistryKey = 'BlankFooter'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtConnTimeOut.Text) & "' where RegistryKey = 'ConnectionTimeOut'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtOrderStatement.Text & "' where RegistryKey = 'OrderStatement'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtStatement.Text & "' where RegistryKey = 'Statement'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkCurrentDateDataEntry.Value & "' where RegistryKey = 'CurrentDateDataEntry'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkIsEntryDate.Value & "' where RegistryKey = 'isEntryDate'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & DtpFrom.DateValue & "' where RegistryKey = 'FromDate'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & DtpTo.DateValue & "' where RegistryKey = 'ToDate'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkIsPortrait.Value & "' where RegistryKey = 'IsPortrait'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkIsLegal.Value & "' where RegistryKey = 'IsLegal'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtEmployeeLateRelaxTime.Text) & "' where RegistryKey = 'EmployeeLateRelaxTime'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtAdminClssingFinePerOnShort.Text) & "' where RegistryKey = 'AdminClssingFinePerOnShort'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtGridRowHeight.Text) & "' where RegistryKey = 'GridRowHeight'")
   'delete CN.Execute ("UPDATE Registry Set Value = '" & ChkAddSpace.Value & "' where RegistryKey = 'AddSpace'")
   
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtPrefixPhoneNo.Text & "' where RegistryKey = 'PrefixPhoneNo'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtOwnerMobileNo.Text & "' where RegistryKey = 'OwnerMobileNo'")
   CN.Execute ("UPDATE sysindexs Set Value = N'" & TxtCustomerSalesMessage.Text & "' where RegistryKey = 'CustomerSalesMessage'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtWebLinkForSMS.Text & "' where RegistryKey = 'WebLinkForSMS'")
   
   CN.Execute ("UPDATE sysindexs Set Value = '" & Val(TxtRoundfigureInSearchForm.Text) & "' where RegistryKey = 'RoundfigureInSearchForm'")
   
   ' Email
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtFromEmail.Text & "' where RegistryKey = 'FromEmail'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtToEmail.Text & "' where RegistryKey = 'ToEmail'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & txtEmailPwd.Text & "' where RegistryKey = 'EmailPwd'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & CmbSMTPServerAddress.Text & "' where RegistryKey = 'SMTPServerAddress'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtPortNo.Text & "' where RegistryKey = 'PortNo'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkUseEmail.Value & "' where RegistryKey = 'UseEmail'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & ChkExportReportASPDF.Value & "' where RegistryKey = 'ExportReportASPDF'")
   CN.Execute ("UPDATE sysindexs Set Value = '" & TxtActivityActionNo.Text & "' where RegistryKey = 'ActivityActionNo'")
   
   MsgBox "Your Software Default Settings has been Changed successfully", vbInformation, "Information"
   'CmbPrinters.Text = !DeviceName & "," & !DriverName & "," & !Port
   ObjRegistry.RefreshRegistry
'   ObjSale.RefreshRegistry
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
  If ChkAllowMonthlyBillNo.Value = 1 And ChkAllowContinuousBillNo.Value = 1 Then
    MsgBox "Please mentioned only one BillNo. either ""Allow Continuous Bill No."" or "" Allow Monthly Bill No.""", vbInformation, "Alert"
    FunValidation = False
    Exit Function
  End If
  'All Ok, now validation is success
  FunValidation = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub SaveLogo()
   On Error GoTo ErrorHandler
   CN.CursorLocation = adUseClient
   CN.Execute "Delete from CompanyLogo"
   strsql = "SELECT * FROM CompanyLogo"
   strFileNm = CD1.FileName
   If strFileNm = "" Then Exit Sub
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open strsql, CN, adOpenStatic, adLockOptimistic
   Rs.AddNew
   DataFile = FreeFile
   Close DataFile
   Open strFileNm For Binary Access Read As DataFile
       Fl = LOF(DataFile)   ' Length of data in file
       If Fl = 0 Then Close DataFile: Exit Sub
       Chunks = Fl \ ChunkSize
       Fragment = Fl Mod ChunkSize
       ReDim Chunk(Fragment)
       Get DataFile, , Chunk()
       Rs!pic.AppendChunk Chunk()
       ReDim Chunk(ChunkSize)
       For i = 1 To Chunks
           Get DataFile, , Chunk()
           Rs!pic.AppendChunk Chunk()
       Next i
   Close DataFile
   Rs.Update
   Rs.Close
   Set Rs = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SaveWaterMark()
   On Error GoTo ErrorHandler
   CN.CursorLocation = adUseClient
   CN.Execute "Delete from WaterMark"
   strsql = "SELECT * FROM WaterMark"
   strFileNm = CD2.FileName
   If strFileNm = "" Then Exit Sub
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open strsql, CN, adOpenStatic, adLockOptimistic
   Rs.AddNew
   DataFile = FreeFile
   Close DataFile
   Open strFileNm For Binary Access Read As DataFile
       Fl = LOF(DataFile)   ' Length of data in file
       If Fl = 0 Then Close DataFile: Exit Sub
       Chunks = Fl \ ChunkSize
       Fragment = Fl Mod ChunkSize
       ReDim Chunk(Fragment)
       Get DataFile, , Chunk()
       Rs!pic.AppendChunk Chunk()
       ReDim Chunk(ChunkSize)
       For i = 1 To Chunks
           Get DataFile, , Chunk()
           Rs!pic.AppendChunk Chunk()
       Next i
   Close DataFile
   Rs.Update
   Rs.Close
   Set Rs = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ShowLogo()
   On Error GoTo ErrorHandler
   strsql = "SELECT * FROM CompanyLogo"
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open strsql, CN, adOpenStatic, adLockOptimistic
   If Rs.RecordCount = 0 Then LogoName = "": Set ImgLogo.Picture = Nothing: Exit Sub
   DataFile = 1
    
   Open "C:\SI.Bmp" For Binary Access Write As DataFile
      Fl = Rs!pic.ActualSize ' Length of data in file
      If Fl = 0 Then Close DataFile: Exit Sub
      Chunks = Fl \ ChunkSize
      Fragment = Fl Mod ChunkSize
      ReDim Chunk(Fragment)
      Chunk() = Rs!pic.GetChunk(Fragment)
      Put DataFile, , Chunk()
      For i = 1 To Chunks
         ReDim Buffer(ChunkSize)
         Chunk() = Rs!pic.GetChunk(ChunkSize)
         Put DataFile, , Chunk()
      Next i
   Close DataFile
   LogoName = "C:\SI.Bmp"
   CD1.FileName = LogoName
   ImgLogo.Picture = LoadPicture(LogoName)
   Rs.Close
   Set Rs = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ShowWaterMark()
   On Error GoTo ErrorHandler
   strsql = "SELECT * FROM WaterMark"
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open strsql, CN, adOpenStatic, adLockOptimistic
   If Rs.RecordCount = 0 Then WaterMarkName = "": Set ImgWaterMark.Picture = Nothing: Exit Sub
   DataFile = 1
    
   Open "C:\WaterMark.Bmp" For Binary Access Write As DataFile
      Fl = Rs!pic.ActualSize ' Length of data in file
      If Fl = 0 Then Close DataFile: Exit Sub
      Chunks = Fl \ ChunkSize
      Fragment = Fl Mod ChunkSize
      ReDim Chunk(Fragment)
      Chunk() = Rs!pic.GetChunk(Fragment)
      Put DataFile, , Chunk()
      For i = 1 To Chunks
         ReDim Buffer(ChunkSize)
         Chunk() = Rs!pic.GetChunk(ChunkSize)
         Put DataFile, , Chunk()
      Next i
   Close DataFile
   LogoName = "C:\WaterMark.Bmp"
   CD2.FileName = LogoName
   ImgWaterMark.Picture = LoadPicture(LogoName)
   Rs.Close
   Set Rs = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hwnd, "Software Default Settings"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   
   PicAllow.Picture = Me.Picture
   PicDefault.Picture = Me.Picture
   PicShowHide.Picture = Me.Picture
   PicSMS.Picture = Me.Picture
   PicEmail.Picture = Me.Picture
      
   vFlag = False
   BtnShowHide_Click
   CmbPrinters.Clear
   CmbPrinters.AddItem "Default,winspool,LPT1"
   Cmb2ndPrinters.Clear
   Cmb2ndPrinters.AddItem "Default,winspool,LPT1"
   Dim p
   For Each p In Printers
      CmbPrinters.AddItem p.DeviceName & "," & p.DriverName & "," & p.Port
      Cmb2ndPrinters.AddItem p.DeviceName & "," & p.DriverName & "," & p.Port
   Next p
   CmbPrinters.ListIndex = 0
   Cmb2ndPrinters.ListIndex = 0
   '''''''''''''''' Form Default Setting  ''''''''''''''''''''''
   sSql = "select * from FormDefaultSetting Where FormType = 'Software Default Setting' and LocalComputerName = '" & LocalComputerName & "'"
   With CN.Execute(sSql)
     If .RecordCount > 0 Then
        If Not IsNull(!DeviceName) Then
            CmbPrinters.Text = !DeviceName & "," & !DriverName & "," & !Port
        Else
            CmbPrinters.ListIndex = 0
        End If
     End If
     .Close
   End With
   ''''''''''''''''''''''''''''''''''''''''''''''
   CmbSMTPServerAddress.Clear
   CmbSMTPServerAddress.AddItem "smtp.live.com"
   CmbSMTPServerAddress.AddItem "smtp.gmail.com"
   CmbSMTPServerAddress.AddItem "smtp.mail.yahoo.com"
   CmbSMTPServerAddress.AddItem "smtp.softinnpk.com"
   CmbSMTPServerAddress.ListIndex = 0
   
   'Allow
   ChkNegativeSale.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'NegativeSale'").Fields(0).Value
   ChkAllowOrderByCodeinInvoices.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowOrderByCodeinInvoices'").Fields(0).Value
   ChkAllowContinuousBillNo.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowContinuousBillNo'").Fields(0).Value
   ChkAllowMonthlyBillNo.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowMonthlyBillNo'").Fields(0).Value
   ChkAllowDailyBillNo.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowDailyBillNo'").Fields(0).Value
   ChkAllowBothPackingsareSame.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowBothPackingsareSame'").Fields(0).Value
   ChkSetEnterKeyGridStockAdjustment.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'SetEnterKeyGridStockAdjustment'").Fields(0).Value
   ChkSaveAsNewBill.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'SaveAsNewBill'").Fields(0).Value
   ChkAfterRowEditFocusNextGridLine.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AfterRowEditFocusNextGridLine'").Fields(0).Value
   ChkHideClearButton.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'HideClearButton'").Fields(0).Value
   ChkAlertAllocateProduct.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AlertAllocateProduct'").Fields(0).Value
   ChkSystemDate.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'SystemDate'").Fields(0).Value
   ChkSaleInProduction.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'SaleInProduction'").Fields(0).Value
   ChkPrintHeadersSaleInvoice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'PrintHeadersSaleInvoice'").Fields(0).Value
   ChkLaserPrintofSaleInvoice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'LaserPrintofSaleInvoice'").Fields(0).Value
   ChkProperCase.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ProperCase'").Fields(0).Value
   ChkDiscountAllowed.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'DiscAllowed'").Fields(0).Value
   ChkCostVisible.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'CostVisible'").Fields(0).Value
   ChkCashReceived.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'CashReceived'").Fields(0).Value
   ChkProductSearchOpenInPreviousState.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ProductSearchOpenInPreviousState'").Fields(0).Value
   ChkChangePrice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ChangePrice'").Fields(0).Value
   ChkAutoApplyPartyLastPrice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AutoApplyPartyLastPrice'").Fields(0).Value
   ChkAutoApplyPartyLastDiscount.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AutoApplyPartyLastDiscount'").Fields(0).Value
   ChkPrintKitchenInoices.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'PrintKitchenInoices'").Fields(0).Value
   ChkAutoPrintSaleOrder.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AutoPrintSaleOrder'").Fields(0).Value
   ChkDisableAutoPrint.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'HideAutoPrint'").Fields(0).Value
   ChkDisableQuantityinPOS.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'DisableQuantityinPOS'").Fields(0).Value
   ChkAllowNegativeStockInBarcodes.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowNegativeStockInBarcodes'").Fields(0).Value
   ChkAllowDiscountOnSaleDistribution.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowDiscountOnSaleDistribution'").Fields(0).Value
   ChkSeperateProductWithPrice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'SeperateProductWithPrice'").Fields(0).Value
   ChkSeperateProductInPOS.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'SeperateProductInPOS'").Fields(0).Value
   ChkAutoEnterQtyintoGridSaleInvoice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AutoEnterQtyintoGridSaleInvoice'").Fields(0).Value
   ChkAutoEnterBeforeQty.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AutoEnterBeforeQty'").Fields(0).Value
   ChkAllowSMSOnSave.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowSMSOnSave'").Fields(0).Value
   ChkAllowSMSOnDelete.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowSMSOnDelete'").Fields(0).Value
   ChkAllowSMSOnClear.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowSMSOnClear'").Fields(0).Value
   ChkAllowSMSOnLogin.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowSMSOnLogin'").Fields(0).Value
   ChkAllowSMSWithDetail.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowSMSWithDetail'").Fields(0).Value
   ChkAllowSMSThroughDevice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowSMSThroughDevice'").Fields(0).Value
   ChkChangeQtyOnChangedPrice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ChangeQtyOnChangedPrice'").Fields(0).Value
   ChkHeaderInfoNotClear.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'HeaderInfoNotClear'").Fields(0).Value
   ChkAutoMoveGridWhenSerialEntered.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AutoMoveGridWhenSerialEntered'").Fields(0).Value
   ChkSerialCompulsoryinInvoice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'SerialCompulsoryinInvoice'").Fields(0).Value
   ChkEmployeeMandatory.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'EmployeeMandatory'").Fields(0).Value
   ChkTableIDMandatory.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'TableIDMandatory'").Fields(0).Value
   ChkUpdateStockSaleBodyInsert.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'UpdateStockSaleBodyInsert'").Fields(0).Value
   ChkEmployeeCommision.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'EmployeeCommision'").Fields(0).Value
   ChkSalePriceLessThanPurchase.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'SalePriceLessThanPurchase'").Fields(0).Value
   ChkChangeQtyPack.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ChangeQtyPack'").Fields(0).Value
   ChkEitherPackORLooseEnter.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'EitherPackORLooseEnter'").Fields(0).Value
   ChkIsSingleBarcode.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'IsSingleBarcode'").Fields(0).Value
   ChkSectorCompulsory.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'SectorCompulsory'").Fields(0).Value
   ChkRemarksCompulsory.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'RemarksCompulsory'").Fields(0).Value
   ChkDivideRetailWithPacking.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'DivideRetailWithPacking'").Fields(0).Value
   ChkUseMultipleStore.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'UseMultipleStore'").Fields(0).Value
   ChkPLSamePR.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'PLSamePR'").Fields(0).Value
   ChkAdminClosingSaveWhenUserClosingSaved.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AdminClosingSaveWhenUserClosingSaved'").Fields(0).Value
   ChkCheckStockOnSave.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'CheckStockOnSave'").Fields(0).Value
   ChkLockPurPrice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'LockPurPrice'").Fields(0).Value
   ChkAutoPrintinInvoices.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AutoPrintinInvoices'").Fields(0).Value
   ChkUsePurPrice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'UsePurPrice'").Fields(0).Value
   ChkChangeTransactionDate.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ChangeTransactionDate'").Fields(0).Value
   ChkUseBin.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'UseBin'").Fields(0).Value
   ChkAttendanceNextDayOut.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AttendanceNextDayOut'").Fields(0).Value
   
   'Show Hide
   ChkBatchNoVisible.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'BatchNoVisible'").Fields(0).Value
   ChkQuantityinBarcodes.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'QuantityinBarcodes'").Fields(0).Value
   ChkShowLastInvoiceMsgAtSave.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowLastInvoiceMsgAtSave'").Fields(0).Value
   ChkTag.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'Tag'").Fields(0).Value
   ChkStoreVisible.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'StoreVisible'").Fields(0).Value
   ChkManualBillNoVisible.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ManualBillNoVisible'").Fields(0).Value
   ChkTableVisible.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'TableVisible'").Fields(0).Value
   ChkOrganizationVisible.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'OrganizationVisible'").Fields(0).Value
   ChkRemarksVisible.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'RemarksVisible'").Fields(0).Value
   ChkMemberVisible.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'MemberVisible'").Fields(0).Value
   ChkSaleOrderVisible.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'SaleOrderVisible'").Fields(0).Value
   ChkFright.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'FreightVisible'").Fields(0).Value
   ChkEmployeeVisible.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'EmpVisible'").Fields(0).Value
   ChkPreviousBalanceVisible.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'PreviousBalanceVisible'").Fields(0).Value
   ChkHideSaleAmount.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'HideSaleAmount'").Fields(0).Value
   ChkHidePurchaseAmount.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'HidePurchaseAmount'").Fields(0).Value
   ChkShowBankInTransection.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowBankInTransection'").Fields(0).Value
   ChkShowWarrantyinSaleInvoice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowWarrantyinSaleInvoice'").Fields(0).Value
   ChkShowCodeInHalfPrint.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowCode'").Fields(0).Value
   ChkShowSerialInHalfPrint.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowSerial'").Fields(0).Value
   ChkShowRetailinPurchaseReturnPrint.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowRetailinPurchaseReturnPrint'").Fields(0).Value
   ChkShowRawMaterialProductInSaleInvoices.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowRawMaterialProductInSaleInvoices'").Fields(0).Value
   ChkInvType.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'InvTypeVisible'").Fields(0).Value
   ChkShowWholeSaleMargin.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowWholeSaleMargin'").Fields(0).Value
   ChkShowPromiseDateInSalaPurchase.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowPromiseDateInSalaPurchase'").Fields(0).Value
   ChkShowSyllabus.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowSyllabus'").Fields(0).Value
   ChkShowProdProfit.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowProdProfit'").Fields(0).Value
   ChkShowStockFromTableGridDataMovement.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowStockFromTableGridDataMovement'").Fields(0).Value
   ChkShowBarcodeProductSearch.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowBarcodeProductSearch'").Fields(0).Value
   ChkOrganizationMandatory.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'OrganizationMandatory'").Fields(0).Value
   ChkisShowPublisher.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'isShowPublisher'").Fields(0).Value
   chkisShowListPrice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'isShowListPrice'").Fields(0).Value
   ChkisShowDepartment.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'isShowDepartment'").Fields(0).Value
   ChkisShowSubDepartment.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'isShowSubDepartment'").Fields(0).Value
   chkisShowSeason.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'isShowSeason'").Fields(0).Value
   ChkisShowItemDesc.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'isShowItemDesc'").Fields(0).Value
   ChkisShowOther.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'isShowOther'").Fields(0).Value
   ChkisShowVendor.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'isShowVendor'").Fields(0).Value
   ChkShowColourSize.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowColourSize'").Fields(0).Value
   ChkIsGrossQty.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'IsGrossQty'").Fields(0).Value
   ChkShowSavedStock.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowSavedStock'").Fields(0).Value
   ChkShowAllPrices.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowAllPrices'").Fields(0).Value
   ChkShowChangeRetailinPurchaseInvoice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowChangeRetailinPurchaseInvoice'").Fields(0).Value
   ChkShowSession.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowSession'").Fields(0).Value
   ChkShowTradeOffer.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowTradeOffer'").Fields(0).Value
   ChkShowExpiryInvoice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowExpiryInvoice'").Fields(0).Value
   ChkShowBonus.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowBonus'").Fields(0).Value
   ChkShowOffer.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowOffer'").Fields(0).Value
   ChkShowSaleTax.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowSaleTax'").Fields(0).Value
   ChkShowSC.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowSC'").Fields(0).Value
   TxtChargesName.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'ChargesName'").Fields(0).Value
   ChkShowBatchPrint.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowBatchPrint'").Fields(0).Value
   ChkShowBarCodeQty.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowBarCodeQty'").Fields(0).Value
   ChkShowHistoryofAllCustomer.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowHistoryofAllCustomer'").Fields(0).Value
   ChkIsRoundFigure.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'IsRoundFigure'").Fields(0).Value
   ChkShowChangePriceOnSavePI.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowChangePriceOnSavePI'").Fields(0).Value
   ChkShowGrandTotalinSearch.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowGrandTotalinSearch'").Fields(0).Value
   ChkShowDispatchDate.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowDispatchDate'").Fields(0).Value
   ChkShowTimeFilterinReport.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'TimeWiseReport'").Fields(0).Value
   ChkShowReSale.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowReSale'").Fields(0).Value
   ChkShowDiscPurPrice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowDiscPurPrice'").Fields(0).Value
   chkAllowNegativeOrder.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'AllowNegativeOrder'").Fields(0).Value
   ChkShowAddBarCode.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowAddBarCode'").Fields(0).Value
   ChkShowPurPrice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowPurPrice'").Fields(0).Value
   ChkProductSearchWithStore.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ProductSearchWithStore'").Fields(0).Value
   ChkShowAllStoreStock.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowAllStoreStock'").Fields(0).Value
   ChkSearchCodeInGrid.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'SearchCodeInGrid'").Fields(0).Value
   ChkUsePasswordForm.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'UsePasswordForm'").Fields(0).Value
   ChkShowPurchaseProfit.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowPurchaseProfit'").Fields(0).Value
   ChkShowBarcodeDesc.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowBarcodeDesc'").Fields(0).Value
   ChkShowDiscPrice.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowDiscPrice'").Fields(0).Value
   ChkShowMultiBranches.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowMultiBranches'").Fields(0).Value
   
   'Default Values
   TxtOrganizationID.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'OrganizationID'").Fields(0).Value
   FunSelectOrganization ssValidate, True
   TxtStoreID.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'StoreID'").Fields(0).Value
   FunSelectStore ssValidate, True
   TxtBankMachineID.Text = IIf(IsNull(CN.Execute("Select value from sysindexs where RegistryKey = 'BankMachineID'").Fields(0).Value), "", CN.Execute("Select value from sysindexs where RegistryKey = 'BankMachineID'").Fields(0).Value)
   FunSelectBankMachine ssValidate, True
   TxtSessionID.Text = IIf(IsNull(CN.Execute("Select value from sysindexs where RegistryKey = 'SessionID'").Fields(0).Value), "", CN.Execute("Select value from sysindexs where RegistryKey = 'SessionID'").Fields(0).Value)
   FunSelectSession ssValidate, True
   TxtNoofPrints.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'NoofPrints'").Fields(0).Value
   TxtMemberMin.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'MemberMin'").Fields(0).Value
   TxtMemberMax.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'MemberMax'").Fields(0).Value
   TxtProdDesc1.Text = CN.Execute("Select isnull(value,'') as value from sysindexs where RegistryKey = 'ProdDesc1'").Fields(0).Value
   TxtPackingChargesPer.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'PackingChargesPer'").Fields(0).Value
   TxtHourDifference.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'HourDifference'").Fields(0).Value
   TxtBarcodePrefix.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'BarCodePrefix'").Fields(0).Value
   TxtSearchDateDifference.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'SearchDateDifference'").Fields(0).Value
   TxtX.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'X'").Fields(0).Value
   TxtY.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'Y'").Fields(0).Value
   TxtCostX.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'CostX'").Fields(0).Value
   TxtCostY.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'CostY'").Fields(0).Value
   TxtBlankFooter.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'BlankFooter'").Fields(0).Value
   TxtConnTimeOut.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'ConnectionTimeOut'").Fields(0).Value
   TxtStatement.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'Statement'").Fields(0).Value
   TxtOrderStatement.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'OrderStatement'").Fields(0).Value
   ChkShowPromiseDateInSalaPurchase.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ShowPromiseDateInSalaPurchase'").Fields(0).Value
   ChkCurrentDateDataEntry.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'CurrentDateDataEntry'").Fields(0).Value
   ChkIsEntryDate.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'isEntryDate'").Fields(0).Value
   DtpFrom.DateValue = CN.Execute("Select value from sysindexs where RegistryKey = 'FromDate'").Fields(0).Value
   DtpTo.DateValue = CN.Execute("Select value from sysindexs where RegistryKey = 'ToDate'").Fields(0).Value
   ChkIsPortrait.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'IsPortrait'").Fields(0).Value
   ChkIsLegal.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'IsLegal'").Fields(0).Value
'   ChkAddSpace.Value = CN.Execute("Select value from Registry where RegistryKey = 'AddSpace'").Fields(0).Value
   TxtEmployeeLateRelaxTime.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'EmployeeLateRelaxTime'").Fields(0).Value
   TxtOwnerMobileNo.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'OwnerMobileNo'").Fields(0).Value
   TxtPrefixPhoneNo.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'PrefixPhoneNo'").Fields(0).Value
   TxtCustomerSalesMessage.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'CustomerSalesMessage'").Fields(0).Value
   TxtWebLinkForSMS.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'WebLinkForSMS'").Fields(0).Value
   TxtAdminClssingFinePerOnShort.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'AdminClssingFinePerOnShort'").Fields(0).Value
   TxtGridRowHeight.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'GridRowHeight'").Fields(0).Value
   TxtRoundfigureInSearchForm.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'RoundfigureInSearchForm'").Fields(0).Value
   
   
   ' Email
   TxtFromEmail.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'FromEmail'").Fields(0).Value
   TxtToEmail.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'ToEmail'").Fields(0).Value
   txtEmailPwd.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'EmailPwd'").Fields(0).Value
   CmbSMTPServerAddress.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'SMTPServerAddress'").Fields(0).Value
   TxtPortNo.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'PortNo'").Fields(0).Value
   ChkUseEmail.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'UseEmail'").Fields(0).Value
   ChkExportReportASPDF.Value = CN.Execute("Select value from sysindexs where RegistryKey = 'ExportReportASPDF'").Fields(0).Value
   TxtActivityActionNo.Text = CN.Execute("Select value from sysindexs where RegistryKey = 'ActivityActionNo'").Fields(0).Value
   
   
   Call ShowLogo
   Call ShowWaterMark
   
   Dim a, a2 As String
'   a = CN.Execute("Select value from sysindexs where RegistryKey = 'DeviceName'").Fields(0).Value & "," & CN.Execute("Select value from sysindexs where RegistryKey = 'DriverName'").Fields(0).Value & "," & CN.Execute("Select value from sysindexs where RegistryKey = 'Port'").Fields(0).Value
'   CmbPrinters.Text = a
'   a2 = CN.Execute("Select value from sysindexs where RegistryKey = 'DeviceName2'").Fields(0).Value & "," & CN.Execute("Select value from sysindexs where RegistryKey = 'DriverName2'").Fields(0).Value & "," & CN.Execute("Select value from sysindexs where RegistryKey = 'Port2'").Fields(0).Value
'   Cmb2ndPrinters.Text = a2

   Exit Sub
ErrorHandler:
   If Err.Number = 383 Then CmbPrinters.ListIndex = 0: Cmb2ndPrinters.ListIndex = 0: Exit Sub
   Call ShowErrorMessage
   Unload Me
End Sub

Private Sub ImgLogo_Click()
   vFlag = True
   CD1.FileName = ""
   CD1.DialogTitle = "Enter Path to take Company Logo"
'   CD1.InitDir = App.Path
   CD1.Filter = "(Image Files)|*.bmp"
   CD1.ShowOpen
   If CD1.FileName <> "" Then
      ImgLogo.Picture = LoadPicture(CD1.FileName)
   Else
      CD1.FileName = ""
      ImgLogo.Picture = Nothing
   End If
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      TxtBankMachineID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub BtnBankMachine_Click()
   If FunSelectBankMachine(ssButton, False) = True Then
      TxtSessionID.SetFocus
   Else
      TxtBankMachineID.SetFocus
   End If
End Sub

Private Sub ImgWaterMark_Click()
   vFlag = True
   CD2.FileName = ""
   CD2.DialogTitle = "Enter Path to Insert Watermark"
'   CD1.InitDir = App.Path
   CD2.Filter = "(Image Files)|*.bmp"
   CD2.ShowOpen
   If CD2.FileName <> "" Then
      ImgWaterMark.Picture = LoadPicture(CD2.FileName)
   Else
      CD2.FileName = ""
      ImgWaterMark.Picture = Nothing
   End If
End Sub

Private Sub TxtBankMachineID_Change()
   If TxtBankMachineID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtBankMachineID.Name Then Exit Sub
   If TxtBankMachineName.Text <> "" Then TxtBankMachineName.Text = ""
End Sub

Private Sub TxtBankMachineID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtBankMachineID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtBankMachineName.Text <> "" Then Exit Sub
   If Trim(TxtBankMachineID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectBankMachine(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectBankMachine(ssButton, False)
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
   If TxtStoreName.Text <> "" Then Exit Sub
   If Trim(TxtStoreID.Text) = "" Then Exit Sub
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
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtOrganizationName.Text = !OrganizationName
          FunSelectOrganization = True
          .Close
          'If btnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectOrganization = False
          .Close
          TxtOrganizationID.Text = ""
          TxtOrganizationName.Text = ""
          'If btnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

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
      TxtStoreID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub TxtSessionID_Change()
   If TxtSessionID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSessionID.Name Then Exit Sub
   If TxtSessionName.Text <> "" Then TxtSessionName.Text = ""
End Sub

Private Sub TxtSessionID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSessionID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSessionName.Text <> "" Then Exit Sub
   If Trim(TxtSessionID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSession(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSession(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSession(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSession.Show vbModal, Me
        If SchSession.ParaOutSessionID = "" Then FunSelectSession = False: Exit Function
        TxtSessionID.Text = SchSession.ParaOutSessionID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Sessions where SessionID=" & Val(TxtSessionID.Text)
    CN.CursorLocation = adUseClient
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSessionName.Text = !SessionName
          FunSelectSession = True
          .Close
          Exit Function
      Else
          FunSelectSession = False
          .Close
          TxtSessionID.Text = ""
          TxtSessionName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
