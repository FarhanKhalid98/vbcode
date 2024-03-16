VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmSoftwareDefaultSettings 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7080
      Left            =   585
      TabIndex        =   3
      Top             =   810
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   12488
      _Version        =   393216
      Tab             =   2
      TabHeight       =   529
      BackColor       =   12632256
      TabCaption(0)   =   "Show/Hide"
      TabPicture(0)   =   "FrmSoftwareDefaultSettings.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ChkOrganizationVisible"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ChkMemberVisible"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ChkTableVisible"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ChkStoreVisible"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ChkEmployeeVisible"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ChkFright"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ChkHideSaleAmount"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ChkHidePurchaseAmount"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ChkTag"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ChkPreviousBalanceVisible"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ChkManualBillNoVisible"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ChkRemarksVisible"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Allow"
      TabPicture(1)   =   "FrmSoftwareDefaultSettings.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ChkSaleInProduction"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ChkDiscountAllowed"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ChkProductSearchOpenInPreviousState"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ChkProperCase"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ChkSystemDate"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ChkAutoApplyPartyLastPrice"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "ChkAddSpace"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "ChkNegativeSale"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "ChkCashReceived"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "ChkCostVisible"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "ChkChangePrice"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "ChkPrintHeadersSaleInvoice"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "ChkLaserPrintofSaleInvoice"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Defaults"
      TabPicture(2)   =   "FrmSoftwareDefaultSettings.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label14"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label8"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label7"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label4"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label13"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label12"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label10"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "ImgLogo"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label11"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label9"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label5"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label20"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label18"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label17"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label16"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label15"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label3"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label2"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label6"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "LblStoreID"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "BtnOrganization"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "TxtOrganizationName"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "TxtOrganizationID"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "BtnBankMachine"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "TxtBankMachineName"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "TxtBankMachineID"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "BtnStore"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "TxtStoreName"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "TxtStoreID"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "TxtProdDesc1"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "TxtHourDifference"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "TxtMemberMax"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "TxtMemberMin"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "TxtNoofPrints"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "TxtY"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "TxtX"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "CmbPrinters"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "TxtOrderStatement"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "TxtStatement"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "TxtSearchDateDifference"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "TxtBlankFooter"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).ControlCount=   41
      Begin VB.CheckBox ChkLaserPrintofSaleInvoice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Laser Print of Sale Invoice Half"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74505
         TabIndex        =   68
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CheckBox ChkPrintHeadersSaleInvoice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print Headers in Sale Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74505
         TabIndex        =   67
         Top             =   1995
         Width           =   2400
      End
      Begin VB.TextBox TxtBlankFooter 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3105
         MaxLength       =   3
         TabIndex        =   50
         Top             =   2925
         Width           =   345
      End
      Begin VB.TextBox TxtSearchDateDifference 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2250
         MaxLength       =   3
         TabIndex        =   48
         Top             =   2565
         Width           =   345
      End
      Begin VB.TextBox TxtStatement 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   45
         MaxLength       =   100
         TabIndex        =   43
         Top             =   810
         Width           =   10515
      End
      Begin VB.TextBox TxtOrderStatement 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   45
         MaxLength       =   100
         TabIndex        =   42
         Top             =   1515
         Width           =   10515
      End
      Begin VB.ComboBox CmbPrinters 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   3015
         Width           =   3315
      End
      Begin VB.TextBox TxtX 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2700
         MaxLength       =   100
         TabIndex        =   36
         Top             =   4140
         Width           =   435
      End
      Begin VB.TextBox TxtY 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3195
         MaxLength       =   100
         TabIndex        =   35
         Top             =   4140
         Width           =   435
      End
      Begin VB.TextBox TxtNoofPrints 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   100
         TabIndex        =   30
         Top             =   4995
         Width           =   885
      End
      Begin VB.TextBox TxtMemberMin 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2700
         MaxLength       =   100
         TabIndex        =   29
         Top             =   4995
         Width           =   885
      End
      Begin VB.TextBox TxtMemberMax 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   4170
         MaxLength       =   100
         TabIndex        =   28
         Top             =   4995
         Width           =   885
      End
      Begin VB.TextBox TxtHourDifference 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   4260
         MaxLength       =   100
         TabIndex        =   27
         Top             =   4230
         Width           =   885
      End
      Begin VB.CheckBox ChkChangePrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Administrator can change Price in Sale Transections"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74505
         TabIndex        =   26
         Top             =   5580
         Width           =   4020
      End
      Begin VB.CheckBox ChkCostVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable Cost Function Keys in Sale Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74505
         TabIndex        =   25
         Top             =   3990
         Width           =   3300
      End
      Begin VB.CheckBox ChkCashReceived 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alllow Auto Cash Received in Sale Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74505
         TabIndex        =   24
         Top             =   4380
         Width           =   3345
      End
      Begin VB.CheckBox ChkNegativeSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Negative Sales"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74505
         TabIndex        =   23
         Top             =   810
         Width           =   1860
      End
      Begin VB.CheckBox ChkAddSpace 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Extra Space at the end of Sale Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74505
         TabIndex        =   22
         Top             =   4785
         Width           =   3390
      End
      Begin VB.CheckBox ChkAutoApplyPartyLastPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Apply Party Last Price in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74505
         TabIndex        =   21
         Top             =   3585
         Width           =   3030
      End
      Begin VB.CheckBox ChkSystemDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Set System Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74505
         TabIndex        =   20
         Top             =   1200
         Width           =   1950
      End
      Begin VB.CheckBox ChkProperCase 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Product Name in Proper Case"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74505
         TabIndex        =   19
         Top             =   2790
         Width           =   2805
      End
      Begin VB.CheckBox ChkProductSearchOpenInPreviousState 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Product Search Open in Previous State"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74505
         TabIndex        =   18
         Top             =   5175
         Width           =   3570
      End
      Begin VB.CheckBox ChkDiscountAllowed 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Discount Allowed For Standard User"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74505
         TabIndex        =   17
         Top             =   3195
         Width           =   2895
      End
      Begin VB.CheckBox ChkSaleInProduction 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Sales In Production"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74505
         TabIndex        =   16
         Top             =   1605
         Width           =   2130
      End
      Begin VB.CheckBox ChkRemarksVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Remarks in Sale Invoice"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74325
         TabIndex        =   15
         Top             =   3195
         Width           =   2490
      End
      Begin VB.CheckBox ChkManualBillNoVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Manual Bill in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74325
         TabIndex        =   14
         Top             =   1845
         Width           =   2355
      End
      Begin VB.CheckBox ChkPreviousBalanceVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Previous Balance in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74325
         TabIndex        =   13
         Top             =   4545
         Width           =   2805
      End
      Begin VB.CheckBox ChkTag 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Tag in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74325
         TabIndex        =   12
         Top             =   945
         Width           =   1860
      End
      Begin VB.CheckBox ChkHidePurchaseAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hide Amount in Purchase Transections For Standard User"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74325
         TabIndex        =   11
         Top             =   5445
         Width           =   4470
      End
      Begin VB.CheckBox ChkHideSaleAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hide Amount in Previous Sale Transections For Standard User"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74325
         TabIndex        =   10
         Top             =   5895
         Width           =   4740
      End
      Begin VB.CheckBox ChkFright 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Freight Option in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74325
         TabIndex        =   9
         Top             =   3645
         Width           =   2580
      End
      Begin VB.CheckBox ChkEmployeeVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Employee in Sale Transections"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74325
         TabIndex        =   8
         Top             =   4995
         Width           =   2985
      End
      Begin VB.CheckBox ChkStoreVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Stores in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74325
         TabIndex        =   7
         Top             =   1395
         Width           =   2040
      End
      Begin VB.CheckBox ChkTableVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Tables in Sale Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74325
         TabIndex        =   6
         Top             =   2295
         Width           =   2400
      End
      Begin VB.CheckBox ChkMemberVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Member in Sale Transections"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74325
         TabIndex        =   5
         Top             =   4095
         Width           =   2805
      End
      Begin VB.CheckBox ChkOrganizationVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Organization in Invoices"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74325
         TabIndex        =   4
         Top             =   2745
         Width           =   2445
      End
      Begin SITextBox.Txt TxtProdDesc1 
         Height          =   315
         Left            =   45
         TabIndex        =   46
         Top             =   2175
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
      Begin SITextBox.Txt TxtStoreID 
         Height          =   315
         Left            =   7065
         TabIndex        =   52
         Tag             =   "NC"
         Top             =   5175
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
         Left            =   8220
         TabIndex        =   53
         Tag             =   "NC"
         Top             =   5175
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
         Left            =   7860
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   5175
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
         MICON           =   "FrmSoftwareDefaultSettings.frx":0054
         BC              =   12632256
         FC              =   0
      End
      Begin SITextBox.Txt TxtBankMachineID 
         Height          =   315
         Left            =   7065
         TabIndex        =   55
         Top             =   5850
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
         Left            =   8220
         TabIndex        =   56
         Top             =   5835
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
         Left            =   7860
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   5835
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
         MICON           =   "FrmSoftwareDefaultSettings.frx":0070
         BC              =   12632256
         FC              =   0
      End
      Begin SITextBox.Txt TxtOrganizationID 
         Height          =   315
         Left            =   7065
         TabIndex        =   58
         Top             =   6510
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
         Left            =   8220
         TabIndex        =   59
         Top             =   6510
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
         Left            =   7860
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   6510
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
         MICON           =   "FrmSoftwareDefaultSettings.frx":008C
         BC              =   12632256
         FC              =   0
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
         Left            =   7065
         TabIndex        =   66
         Top             =   4950
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
         Left            =   7065
         TabIndex        =   65
         Top             =   5625
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
         Left            =   8235
         TabIndex        =   64
         Top             =   4950
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
         Left            =   8820
         TabIndex        =   63
         Top             =   5625
         Width           =   1770
      End
      Begin VB.Label Label15 
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
         Left            =   8820
         TabIndex        =   62
         Top             =   6300
         Width           =   1620
      End
      Begin VB.Label Label16 
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
         Left            =   7065
         TabIndex        =   61
         Top             =   6300
         Width           =   1335
      End
      Begin VB.Label Label17 
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
         Left            =   90
         TabIndex        =   51
         Top             =   2970
         Width           =   3000
      End
      Begin VB.Label Label18 
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
         Left            =   135
         TabIndex        =   49
         Top             =   2610
         Width           =   2025
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
         Left            =   75
         TabIndex        =   47
         Top             =   1935
         Width           =   3375
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
         Left            =   45
         TabIndex        =   45
         Top             =   585
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
         Left            =   45
         TabIndex        =   44
         Top             =   1290
         Width           =   1995
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
         Left            =   3600
         TabIndex        =   41
         Top             =   2745
         Width           =   2235
      End
      Begin VB.Image ImgLogo 
         Height          =   645
         Left            =   5790
         Stretch         =   -1  'True
         Top             =   3765
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
         Left            =   5400
         TabIndex        =   39
         Top             =   3465
         Width           =   1260
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
         Left            =   2835
         TabIndex        =   38
         Top             =   3915
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
         Left            =   3330
         TabIndex        =   37
         Top             =   3915
         Width           =   135
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
         Left            =   270
         TabIndex        =   34
         Top             =   4740
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
         Left            =   2700
         TabIndex        =   33
         Top             =   4770
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
         Left            =   4170
         TabIndex        =   32
         Top             =   4770
         Width           =   1335
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
         Left            =   4035
         TabIndex        =   31
         Top             =   3960
         Width           =   1365
      End
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   4695
      TabIndex        =   0
      Top             =   8265
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
      MICON           =   "FrmSoftwareDefaultSettings.frx":00A8
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   6045
      TabIndex        =   1
      Top             =   8265
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
      MICON           =   "FrmSoftwareDefaultSettings.frx":00C4
      BC              =   14737632
      FC              =   0
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   11100
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   1980
      TabIndex        =   2
      Top             =   195
      Width           =   3465
   End
   Begin VB.Image ImgExit 
      Height          =   360
      Left            =   11625
      Top             =   0
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

Private Sub ChkTableVisible_Click()

End Sub

Private Sub CmbPrinters_Change()

End Sub

Private Sub Label11_Click()

End Sub

Private Sub SSTab1_DblClick()

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
'      If TxtCustomerID.Enabled Then TxtCustomerID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

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

Private Sub ChkStoreVisible_Click()
   TxtStoreID.Enabled = ChkStoreVisible.Value = 1
   TxtStoreName.Enabled = TxtStoreID.Enabled
   BtnStore.Enabled = TxtStoreID.Enabled
End Sub

Private Sub ChkOrganizationVisible_Click()
   TxtOrganizationID.Enabled = ChkOrganizationVisible.Value = 1
   TxtOrganizationName.Enabled = TxtOrganizationID.Enabled
   BtnOrganization.Enabled = TxtOrganizationID.Enabled
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtBankMachineID.SetFocus
         Case TxtBankMachineID.Name: If FunSelectBankMachine(ssFunctionKey, False) = True Then TxtOrganizationID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then TxtStatement.SetFocus
      End Select
   End If
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   If FunValidation = False Then Exit Sub
   Dim vPrinter() As String
   vPrinter = Split(CmbPrinters.Text, ",")
   Call SaveLogo
   
   CN.Execute ("UPDATE Registry Set Value = '" & ChkOrganizationVisible.Value & "' where RegistryKey = 'OrganizationVisible'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkStoreVisible.Value & "' where RegistryKey = 'StoreVisible'")
   
   CN.Execute ("UPDATE Registry Set Value = '" & IIf(Trim(TxtStoreID.Text) = "", "Null", Val(TxtStoreID.Text)) & "' where RegistryKey = 'StoreID'")
   CN.Execute ("UPDATE Registry Set Value = '" & IIf(Trim(TxtBankMachineID.Text) = "", "Null", Val(TxtBankMachineID.Text)) & "' where RegistryKey = 'BankMachineID'")
   CN.Execute ("UPDATE Registry Set Value = '" & IIf(Trim(TxtOrganizationID.Text) = "", "Null", Val(TxtOrganizationID.Text)) & "' where RegistryKey = 'OrganizationID'")
   
   CN.Execute ("UPDATE Registry Set Value = '" & ChkAddSpace.Value & "' where RegistryKey = 'AddSpace'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkNegativeSale.Value & "' where RegistryKey = 'NegativeSale'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkCashReceived.Value & "' where RegistryKey = 'CashReceived'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkCostVisible.Value & "' where RegistryKey = 'CostVisible'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkChangePrice.Value & "' where RegistryKey = 'ChangePrice'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkEmployeeVisible.Value & "' where RegistryKey = 'EmpVisible'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkMemberVisible.Value & "' where RegistryKey = 'MemberVisible'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkManualBillNoVisible.Value & "' where RegistryKey = 'ManualBillNoVisible'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkRemarksVisible.Value & "' where RegistryKey = 'RemarksVisible'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkPreviousBalanceVisible.Value & "' where RegistryKey = 'PreviousBalanceVisible'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkHideSaleAmount.Value & "' where RegistryKey = 'HideSaleAmount'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkHidePurchaseAmount.Value & "' where RegistryKey = 'HidePurchaseAmount'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkLaserPrintofSaleInvoice.Value & "' where RegistryKey = 'LaserPrintofSaleInvoice'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkPrintHeadersSaleInvoice.Value & "' where RegistryKey = 'PrintHeadersSaleInvoice'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkRemarksVisible.Value & "' where RegistryKey = 'RemarksVisible'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkTag.Value & "' where RegistryKey = 'Tag'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkFright.Value & "' where RegistryKey = 'FreightVisible'")
   CN.Execute ("UPDATE Registry Set Value = '" & ChkAutoApplyPartyLastPrice.Value & "' where RegistryKey = 'AutoApplyPartyLastPrice'")
      
   CN.Execute ("UPDATE Registry Set Value = '" & Val(TxtNoofPrints.Text) & "' where RegistryKey = 'NoofPrints'")
   CN.Execute ("UPDATE Registry Set Value = '" & Val(TxtHourDifference.Text) & "' where RegistryKey = 'HourDifference'")
   CN.Execute ("UPDATE Registry Set Value = '" & Val(TxtMemberMin.Text) & "' where RegistryKey = 'MemberMin'")
   CN.Execute ("UPDATE Registry Set Value = '" & Val(TxtMemberMax.Text) & "' where RegistryKey = 'MemberMax'")
   CN.Execute ("UPDATE Registry Set Value = '" & Val(TxtX.Text) & "' where RegistryKey = 'X'")
   CN.Execute ("UPDATE Registry Set Value = '" & Val(TxtY.Text) & "' where RegistryKey = 'Y'")
   CN.Execute ("UPDATE Registry Set Value = '" & TxtOrderStatement.Text & "' where RegistryKey = 'OrderStatement'")
   CN.Execute ("UPDATE Registry Set Value = '" & TxtStatement.Text & "' where RegistryKey = 'Statement'")
   CN.Execute ("UPDATE Registry Set Value = '" & vPrinter(0) & "' where RegistryKey = 'DeviceName'")
   CN.Execute ("UPDATE Registry Set Value = '" & vPrinter(1) & "' where RegistryKey = 'DriverName'")
   CN.Execute ("UPDATE Registry Set Value = '" & vPrinter(2) & "' where RegistryKey = 'Port'")
   MsgBox "Your Main Settings has been Changed successfully", vbInformation, "Information"
   'CmbPrinters.Text = !DeviceName & "," & !DriverName & "," & !Port
  ObjRegistry.RefreshRegistry
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
'  If Trim(TxtName.Text) = "" Then
'    MsgBox "Please specify a Company Name", vbExclamation, "Alert"
'    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
'    Exit Function
'  End If
  'All Ok, now validation is success
  FunValidation = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub SaveLogo()
   On Error GoTo ErrorHandler
   CN.Execute "Delete from CompanyLogo"
   strsql = "SELECT * FROM CompanyLogo"
   strFileNm = CD1.FileName
   If strFileNm = "" Then Exit Sub
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open strsql, CN, adOpenStatic, adLockOptimistic
   Rs.AddNew
   DataFile = 1
   Close DataFile
   Open strFileNm For Binary Access Read As DataFile
       Fl = LOF(DataFile)   ' Length of data in file
       If Fl = 0 Then Close DataFile: Exit Sub
       Chunks = Fl \ ChunkSize
       Fragment = Fl Mod ChunkSize
       ReDim Chunk(Fragment)
       Get DataFile, , Chunk()
       Rs!Pic.AppendChunk Chunk()
       ReDim Chunk(ChunkSize)
       For i = 1 To Chunks
           Get DataFile, , Chunk()
           Rs!Pic.AppendChunk Chunk()
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
      Fl = Rs!Pic.ActualSize ' Length of data in file
      If Fl = 0 Then Close DataFile: Exit Sub
      Chunks = Fl \ ChunkSize
      Fragment = Fl Mod ChunkSize
      ReDim Chunk(Fragment)
      Chunk() = Rs!Pic.GetChunk(Fragment)
      Put DataFile, , Chunk()
      For i = 1 To Chunks
         ReDim Buffer(ChunkSize)
         Chunk() = Rs!Pic.GetChunk(ChunkSize)
         Put DataFile, , Chunk()
      Next i
   Close DataFile
   LogoName = "C:\SI.Bmp"
   ImgLogo.Picture = LoadPicture(LogoName)
   Rs.Close
   Set Rs = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "Main Settings"
   ShowPicture Me
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   vFlag = False
   Dim p
   For Each p In Printers
      CmbPrinters.AddItem p.DeviceName & "," & p.DriverName & "," & p.Port
   Next p
     
   ChkOrganizationVisible.Value = CN.Execute("Select value from Registry where RegistryKey = 'OrganizationVisible'").Fields(0).Value
   ChkStoreVisible.Value = CN.Execute("Select value from Registry where RegistryKey = 'StoreVisible'").Fields(0).Value
   TxtStoreID.Text = CN.Execute("Select value from Registry where RegistryKey = 'StoreID'").Fields(0).Value
   FunSelectStore ssValidate, True
   TxtBankMachineID.Text = IIf(IsNull(CN.Execute("Select value from Registry where RegistryKey = 'BankMachineID'").Fields(0).Value), "", CN.Execute("Select value from Registry where RegistryKey = 'BankMachineID'").Fields(0).Value)
   FunSelectBankMachine ssValidate, True
   TxtOrganizationID.Text = CN.Execute("Select value from Registry where RegistryKey = 'OrganizationID'").Fields(0).Value
   FunSelectOrganization ssValidate, True
   ChkAddSpace.Value = CN.Execute("Select value from Registry where RegistryKey = 'AddSpace'").Fields(0).Value
   ChkNegativeSale.Value = CN.Execute("Select value from Registry where RegistryKey = 'NegativeSale'").Fields(0).Value
   ChkCashReceived.Value = CN.Execute("Select value from Registry where RegistryKey = 'CashReceived'").Fields(0).Value
   ChkCostVisible.Value = CN.Execute("Select value from Registry where RegistryKey = 'CostVisible'").Fields(0).Value
   ChkChangePrice.Value = CN.Execute("Select value from Registry where RegistryKey = 'ChangePrice'").Fields(0).Value
   ChkEmployeeVisible.Value = CN.Execute("Select value from Registry where RegistryKey = 'EmpVisible'").Fields(0).Value
   ChkMemberVisible.Value = CN.Execute("Select value from Registry where RegistryKey = 'MemberVisible'").Fields(0).Value
   ChkManualBillNoVisible.Value = CN.Execute("Select value from Registry where RegistryKey = 'ManualBillNoVisible'").Fields(0).Value
   ChkRemarksVisible.Value = CN.Execute("Select value from Registry where RegistryKey = 'RemarksVisible'").Fields(0).Value
   ChkPreviousBalanceVisible.Value = CN.Execute("Select value from Registry where RegistryKey = 'PreviousBalanceVisible'").Fields(0).Value
   ChkHideSaleAmount.Value = CN.Execute("Select value from Registry where RegistryKey = 'HideSaleAmount'").Fields(0).Value
   ChkHidePurchaseAmount.Value = CN.Execute("Select value from Registry where RegistryKey = 'HidePurchaseAmount'").Fields(0).Value
   ChkLaserPrintofSaleInvoice.Value = CN.Execute("Select value from Registry where RegistryKey = 'LaserPrintofSaleInvoice'").Fields(0).Value
   ChkPrintHeadersSaleInvoice.Value = CN.Execute("Select value from Registry where RegistryKey = 'PrintHeadersSaleInvoice'").Fields(0).Value
   ChkTag.Value = CN.Execute("Select value from Registry where RegistryKey = 'Tag'").Fields(0).Value
   ChkFright.Value = CN.Execute("Select value from Registry where RegistryKey = 'FreightVisible'").Fields(0).Value
   ChkAutoApplyPartyLastPrice.Value = CN.Execute("Select value from Registry where RegistryKey = 'AutoApplyPartyLastPrice'").Fields(0).Value
   TxtX.Text = CN.Execute("Select value from Registry where RegistryKey = 'X'").Fields(0).Value
   TxtY.Text = CN.Execute("Select value from Registry where RegistryKey = 'Y'").Fields(0).Value
   TxtHourDifference.Text = CN.Execute("Select value from Registry where RegistryKey = 'HourDifference'").Fields(0).Value
   TxtNoofPrints.Text = CN.Execute("Select value from Registry where RegistryKey = 'NoofPrints'").Fields(0).Value
   TxtMemberMax.Text = CN.Execute("Select value from Registry where RegistryKey = 'MemberMax'").Fields(0).Value
   TxtMemberMin.Text = CN.Execute("Select value from Registry where RegistryKey = 'MemberMin'").Fields(0).Value
   TxtStatement.Text = CN.Execute("Select value from Registry where RegistryKey = 'Statement'").Fields(0).Value
   TxtOrderStatement.Text = CN.Execute("Select value from Registry where RegistryKey = 'OrderStatement'").Fields(0).Value
   
   Call ShowLogo

   Dim a As String
   a = CN.Execute("Select value from Registry where RegistryKey = 'DeviceName'").Fields(0).Value & "," & CN.Execute("Select value from Registry where RegistryKey = 'DriverName'").Fields(0).Value & "," & CN.Execute("Select value from Registry where RegistryKey = 'Port'").Fields(0).Value
   CmbPrinters.Text = CN.Execute("Select value from Registry where RegistryKey = 'DeviceName'").Fields(0).Value & "," & CN.Execute("Select value from Registry where RegistryKey = 'DriverName'").Fields(0).Value & "," & CN.Execute("Select value from Registry where RegistryKey = 'Port'").Fields(0).Value
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
      TxtStatement.SetFocus
   Else
      TxtBankMachineID.SetFocus
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
