VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmMainSettings 
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
   Begin VB.CheckBox ChkAutoApplyPartyLastPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Auto Apply Party Last Price in Invoices"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   13
      Top             =   6300
      Width           =   3030
   End
   Begin VB.CheckBox ChkFright 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Freight Option in Invoices"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3015
      TabIndex        =   12
      Top             =   5850
      Width           =   2580
   End
   Begin VB.CheckBox ChkOrganizationVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Organization in Invoices"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   0
      Top             =   1080
      Width           =   2580
   End
   Begin VB.TextBox TxtHourDifference 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   9990
      MaxLength       =   100
      TabIndex        =   51
      Top             =   4725
      Width           =   885
   End
   Begin VB.CheckBox ChkStoreVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Stores in Invoices"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   1
      Top             =   1485
      Width           =   2085
   End
   Begin VB.CheckBox ChkAddSpace 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Extra Space at the end of Sale Invoice"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   2
      Top             =   2775
      Width           =   3795
   End
   Begin VB.CheckBox ChkNegativeSale 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Allow Negative Sales"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   3
      Top             =   3225
      Width           =   1905
   End
   Begin VB.CheckBox ChkCashReceived 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Alllow Auto Cash Received in Sale Invoice"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   4
      Top             =   3660
      Width           =   3435
   End
   Begin VB.CheckBox ChkCostVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enable Cost Function Keys in Sale Invoice"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   5
      Top             =   4095
      Width           =   3390
   End
   Begin VB.CheckBox ChkChangePrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Administrator can change Price in Sale Transections"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   6
      Top             =   4545
      Width           =   4065
   End
   Begin VB.CheckBox ChkEmployeeVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Employee in Sale Transections"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   765
      TabIndex        =   7
      Top             =   1875
      Width           =   2985
   End
   Begin VB.CheckBox ChkMemberVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Member in Sale Transections"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   765
      TabIndex        =   8
      Top             =   2310
      Width           =   2850
   End
   Begin VB.CheckBox ChkHideSaleAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hide Amount in Previous Sale Transections For Standard User"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   9
      Top             =   5010
      Width           =   4740
   End
   Begin VB.CheckBox ChkHidePurchaseAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hide Amount in Purchase Transections For Standard User"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   10
      Top             =   5445
      Width           =   4470
   End
   Begin VB.CheckBox ChkTag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Tag in Invoices"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   810
      TabIndex        =   11
      Top             =   5850
      Width           =   2085
   End
   Begin VB.TextBox TxtY 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   9945
      MaxLength       =   100
      TabIndex        =   20
      Top             =   3645
      Width           =   435
   End
   Begin VB.TextBox TxtX 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   9450
      MaxLength       =   100
      TabIndex        =   19
      Top             =   3645
      Width           =   435
   End
   Begin VB.CheckBox ChkPreviousBalanceVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Previous Balance in Invoices"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6705
      TabIndex        =   23
      Top             =   4815
      Width           =   2850
   End
   Begin VB.CheckBox ChkManualBillNoVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Manual Bill in Invoices"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6705
      TabIndex        =   22
      Top             =   4455
      Width           =   2445
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      Left            =   6735
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   6345
      Width           =   3315
   End
   Begin VB.TextBox TxtOrderStatement 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   810
      MaxLength       =   100
      TabIndex        =   29
      Top             =   7635
      Width           =   10515
   End
   Begin VB.CheckBox ChkRemarksVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Remarks in Sale Invoice"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6705
      TabIndex        =   21
      Top             =   4110
      Width           =   2670
   End
   Begin VB.TextBox TxtMemberMax 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   9900
      MaxLength       =   100
      TabIndex        =   26
      Top             =   5490
      Width           =   885
   End
   Begin VB.TextBox TxtMemberMin 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   8430
      MaxLength       =   100
      TabIndex        =   25
      Top             =   5490
      Width           =   885
   End
   Begin VB.TextBox TxtNoofPrints 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   6735
      MaxLength       =   100
      TabIndex        =   24
      Top             =   5490
      Width           =   885
   End
   Begin VB.CheckBox ChkPrintHeadersSaleInvoice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Print Headers in Sale Invoice"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6705
      TabIndex        =   18
      Top             =   3750
      Width           =   2535
   End
   Begin VB.CheckBox ChkLaserPrintofSaleInvoice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Laser Print of Sale Invoice Half"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6705
      TabIndex        =   17
      Top             =   3420
      Width           =   2580
   End
   Begin VB.TextBox TxtStatement 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   810
      MaxLength       =   100
      TabIndex        =   28
      Top             =   6930
      Width           =   10515
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   4695
      TabIndex        =   30
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
      MICON           =   "FrmMainSettings.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   6045
      TabIndex        =   31
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
      MICON           =   "FrmMainSettings.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   6795
      TabIndex        =   14
      Tag             =   "NC"
      Top             =   1530
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
      Left            =   7950
      TabIndex        =   34
      Tag             =   "NC"
      Top             =   1530
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
      Left            =   7590
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1530
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
      MICON           =   "FrmMainSettings.frx":0038
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtBankMachineID 
      Height          =   315
      Left            =   6795
      TabIndex        =   15
      Top             =   2190
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
      Left            =   7950
      TabIndex        =   37
      Top             =   2190
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
      Left            =   7590
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2190
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
      MICON           =   "FrmMainSettings.frx":0054
      BC              =   12632256
      FC              =   0
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   11100
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   6795
      TabIndex        =   16
      Top             =   2865
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
      Left            =   7950
      TabIndex        =   53
      Top             =   2865
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
      Left            =   7590
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   2865
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
      MICON           =   "FrmMainSettings.frx":0070
      BC              =   12632256
      FC              =   0
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
      Left            =   6795
      TabIndex        =   56
      Top             =   2655
      Width           =   1335
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
      Left            =   8550
      TabIndex        =   55
      Top             =   2655
      Width           =   1620
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
      Left            =   9765
      TabIndex        =   52
      Top             =   4455
      Width           =   1365
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
      Left            =   10080
      TabIndex        =   50
      Top             =   3420
      Width           =   135
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
      Left            =   9585
      TabIndex        =   49
      Top             =   3420
      Width           =   135
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
      Left            =   6735
      TabIndex        =   48
      Top             =   6075
      Width           =   2235
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
      Left            =   10740
      TabIndex        =   47
      Top             =   3390
      Width           =   1260
   End
   Begin VB.Image ImgLogo 
      Height          =   645
      Left            =   11130
      Stretch         =   -1  'True
      Top             =   3690
      Width           =   585
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
      Left            =   810
      TabIndex        =   46
      Top             =   7410
      Width           =   1995
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
      Left            =   9900
      TabIndex        =   45
      Top             =   5265
      Width           =   1335
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
      Left            =   8430
      TabIndex        =   44
      Top             =   5265
      Width           =   1290
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
      Left            =   6000
      TabIndex        =   43
      Top             =   5235
      Width           =   2355
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
      Left            =   8550
      TabIndex        =   42
      Top             =   1980
      Width           =   1770
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
      Left            =   7965
      TabIndex        =   41
      Top             =   1305
      Width           =   1005
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   435
      X2              =   705
      Y1              =   78
      Y2              =   78
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2610
      Left            =   6525
      Top             =   675
      Width           =   4065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Box Setting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6570
      TabIndex        =   40
      Top             =   810
      Width           =   3960
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
      Left            =   6795
      TabIndex        =   39
      Top             =   1980
      Width           =   1485
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
      Left            =   6795
      TabIndex        =   36
      Top             =   1305
      Width           =   720
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Settings"
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
      TabIndex        =   33
      Top             =   195
      Width           =   1890
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
      Left            =   810
      TabIndex        =   32
      Top             =   6705
      Width           =   1785
   End
   Begin VB.Image ImgExit 
      Height          =   360
      Left            =   11625
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "FrmMainSettings"
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
