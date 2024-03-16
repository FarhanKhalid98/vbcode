VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptProductionRegister 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptProductionRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   8295
      TabIndex        =   48
      Top             =   4275
      Width           =   4215
      Begin VB.OptionButton RdoBoth 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Both"
         Height          =   255
         Left            =   3060
         TabIndex        =   9
         Top             =   60
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton RdoOut 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Production Out"
         Height          =   255
         Left            =   1575
         TabIndex        =   8
         Top             =   60
         Width           =   1455
      End
      Begin VB.OptionButton RdoIn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Production In"
         Height          =   255
         Left            =   270
         TabIndex        =   7
         Top             =   60
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   10245
      TabIndex        =   47
      Top             =   4875
      Width           =   2250
      Begin VB.OptionButton RdoSummary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Summary"
         Height          =   255
         Left            =   975
         TabIndex        =   12
         Top             =   10
         Width           =   960
      End
      Begin VB.OptionButton RdoDetail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detail"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   10
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   8295
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4875
      Width           =   1950
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   11145
      TabIndex        =   17
      Top             =   6585
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
      MICON           =   "RptProductionRegister.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8370
      TabIndex        =   15
      Top             =   6585
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
      MICON           =   "RptProductionRegister.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9750
      TabIndex        =   16
      Top             =   6585
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
      MICON           =   "RptProductionRegister.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   6785
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
      IntegralPoint   =   15
   End
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   3660
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6785
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
      MICON           =   "RptProductionRegister.frx":0F1E
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   4020
      TabIndex        =   28
      Top             =   6785
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
   Begin JeweledBut.JeweledButton BtnFromStore 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3660
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3585
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
      MICON           =   "RptProductionRegister.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3660
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5505
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
      MICON           =   "RptProductionRegister.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   5505
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
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   4020
      TabIndex        =   24
      Tag             =   "nc"
      Top             =   5505
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
   Begin JeweledBut.JeweledButton BtnSubGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3660
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   6145
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
      MICON           =   "RptProductionRegister.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubGroupID 
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   6145
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
   Begin SITextBox.Txt TxtSubGroupName 
      Height          =   315
      Left            =   4020
      TabIndex        =   26
      Tag             =   "nc"
      Top             =   6145
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
   Begin JeweledBut.JeweledButton BtnCompany 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3660
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4865
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
      MICON           =   "RptProductionRegister.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   4865
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
   Begin SITextBox.Txt TxtCompanyName 
      Height          =   315
      Left            =   4020
      TabIndex        =   23
      Tag             =   "nc"
      Top             =   4865
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
   Begin JeweledBut.JeweledButton BtnUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3660
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   7425
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
      MICON           =   "RptProductionRegister.frx":0FAA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtUserNo 
      Height          =   315
      Left            =   2640
      TabIndex        =   6
      Top             =   7425
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
   Begin SITextBox.Txt TxtUserName 
      Height          =   315
      Left            =   4020
      TabIndex        =   21
      Tag             =   "nc"
      Top             =   7425
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
   Begin SITextBox.Txt TxtFromStoreID 
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   3585
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
   Begin SITextBox.Txt TxtFromStoreName 
      Height          =   315
      Left            =   4020
      TabIndex        =   19
      Tag             =   "nc"
      Top             =   3585
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
   Begin JeweledBut.JeweledButton BtnToStore 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3675
      TabIndex        =   45
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4225
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
      MICON           =   "RptProductionRegister.frx":0FC6
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtToStoreID 
      Height          =   315
      Left            =   2655
      TabIndex        =   1
      Top             =   4225
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
   Begin SITextBox.Txt TxtToStoreName 
      Height          =   315
      Left            =   4035
      TabIndex        =   46
      Tag             =   "nc"
      Top             =   4225
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   8865
      TabIndex        =   13
      Top             =   5625
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
      Left            =   10620
      TabIndex        =   14
      Top             =   5625
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
   Begin VB.Label Label5 
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
      Left            =   8865
      TabIndex        =   50
      Top             =   5400
      Width           =   885
   End
   Begin VB.Label Label6 
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
      Left            =   10635
      TabIndex        =   49
      Top             =   5400
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   541
      X2              =   541
      Y1              =   275
      Y2              =   411
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   541
      X2              =   845
      Y1              =   275
      Y2              =   275
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   847
      X2              =   847
      Y1              =   275
      Y2              =   411
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   845
      X2              =   541
      Y1              =   411
      Y2              =   411
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Store ID"
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
      Left            =   2655
      TabIndex        =   44
      Top             =   4020
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Store Name"
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
      Left            =   4020
      TabIndex        =   43
      Top             =   4020
      Width           =   1290
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Store Name"
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
      Left            =   4020
      TabIndex        =   42
      Top             =   3390
      Width           =   1470
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Store ID"
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
      Left            =   2640
      TabIndex        =   41
      Top             =   3390
      Width           =   1185
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Left            =   4020
      TabIndex        =   40
      Top             =   7230
      Width           =   945
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
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
      Left            =   2640
      TabIndex        =   39
      Top             =   7230
      Width           =   660
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
      Left            =   2640
      TabIndex        =   38
      Top             =   5295
      Width           =   780
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
      Left            =   4020
      TabIndex        =   37
      Top             =   5295
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
      Left            =   2640
      TabIndex        =   36
      Top             =   5925
      Width           =   1170
   End
   Begin VB.Label Label12 
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
      Left            =   2640
      TabIndex        =   35
      Top             =   4650
      Width           =   1035
   End
   Begin VB.Label Label11 
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
      Left            =   4020
      TabIndex        =   34
      Top             =   4650
      Width           =   1320
   End
   Begin VB.Label Label3 
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
      Left            =   4020
      TabIndex        =   33
      Top             =   5925
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
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
      Height          =   195
      Left            =   2640
      TabIndex        =   32
      Top             =   6555
      Width           =   930
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
      Left            =   4020
      TabIndex        =   31
      Top             =   6555
      Width           =   1215
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Production Register"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2700
      TabIndex        =   30
      Top             =   270
      Width           =   2295
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
Attribute VB_Name = "RptProductionRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim sSQL As String

Private Sub BtnCompany_Click()
   If FunSelectCompany(ssButton, False) = True Then
      TxtGroupID.SetFocus
   Else
      TxtCompanyID.SetFocus
   End If
End Sub

Private Sub BtnFromStore_Click()
   If FunSelectToStore(ssButton, False) = True Then
     TxtToStoreID.SetFocus
   Else
      TxtFromStoreID.SetFocus
   End If
End Sub

Private Sub BtnGroup_Click()
   If FunSelectGroup(ssButton, False) = True Then
      TxtSubGroupID.SetFocus
   Else
      TxtGroupID.SetFocus
   End If
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      TxtUserNo.SetFocus
   Else
      TxtProductID.SetFocus
   End If
End Sub

Private Sub BtnSubGroup_Click()
   If FunSelectSubGroup(ssButton, False) = True Then
      TxtProductID.SetFocus
   Else
      TxtSubGroupID.SetFocus
   End If
End Sub

Private Sub BtnToStore_Click()
   If FunSelectToStore(ssButton, False) = True Then
     TxtCompanyID.SetFocus
   Else
      TxtToStoreID.SetFocus
   End If
End Sub

Private Sub BtnUser_Click()
   If FunSelectUser(ssButton, False) = True Then
      RdoIn.SetFocus
   Else
      TxtUserNo.SetFocus
   End If
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

Private Function FunSelectFromStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectFromStore = False: Exit Function
        TxtFromStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    If Trim(TxtFromStoreID.Text) = "" Then Exit Function
    If Len(TxtFromStoreID.Text) <= 3 Then
      TxtFromStoreID.Text = Right("000" + CStr(Val(TxtFromStoreID.Text)), 3)
    End If
    If TxtFromStoreID.Text = "" Then FunSelectFromStore = False: Exit Function
    vStrSQL = " Select StoreName FROM Stores where StoreID='" & TxtFromStoreID.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtFromStoreName.Text = !StoreName
          FunSelectFromStore = True
          .Close
          Exit Function
      Else
          FunSelectFromStore = False
          .Close
          TxtFromStoreID.Text = ""
          TxtFromStoreName.Text = ""
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
    vStrSQL = " Select * FROM Groups where GroupID = '" & TxtGroupID.Text & "'"
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

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtProductID.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtProductID.Text) = "" Then Exit Function
    If Len(TxtProductID.Text) <= 5 Then
      TxtProductID.Text = Right("00000" + CStr(Val(TxtProductID.Text)), 5)
    End If
    If TxtProductID.Text = "" Then FunSelectProduct = False: Exit Function
    vStrSQL = " SELECT p.productid, code, ProductName" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where p.productid = '" & TxtProductID.Text & "' or code='" & TxtProductID.Text & "'"
  
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         'TxtProductID.Text = !ProductID
         TxtProductName.Text = !ProductName
         FunSelectProduct = True
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtProductID.Text = ""
         TxtProductName.Text = ""
         Exit Function
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
    vStrSQL = " Select * FROM SubGroups where SubGroupID = " & Val(TxtSubGroupID.Text)
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

Private Function FunSelectToStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectToStore = False: Exit Function
        TxtToStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    If Trim(TxtToStoreID.Text) = "" Then Exit Function
    If Len(TxtToStoreID.Text) <= 3 Then
      TxtToStoreID.Text = Right("000" + CStr(Val(TxtToStoreID.Text)), 3)
    End If
    If TxtToStoreID.Text = "" Then FunSelectToStore = False: Exit Function
    vStrSQL = " Select StoreName FROM Stores where StoreID='" & TxtToStoreID.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtToStoreName.Text = !StoreName
          FunSelectToStore = True
          .Close
          Exit Function
      Else
          FunSelectToStore = False
          .Close
          TxtToStoreID.Text = ""
          TxtToStoreName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectUser(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchUser.Show vbModal, Me
        If SchUser.ParaOutUserNo = "" Then FunSelectUser = False: Exit Function
        TxtUserNo.Text = SchUser.ParaOutUserNo
    End If
    '---------------------------
    vStrSQL = " Select * FROM Users where UserNo=" & Val(TxtUserNo.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtUserName.Text = !UserName
          FunSelectUser = True
          .Close
          Exit Function
      Else
          FunSelectUser = False
          .Close
          TxtUserNo.Text = ""
          TxtUserName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
      If RdoDetail.Value = True Then
        If RdoIn.Value = True Then
            RptReportViewer.Caption = "Production In Detail (" & CmbGroup.Text & ")"
        ElseIf RdoOut.Value = True Then
            RptReportViewer.Caption = "Production Out Detail (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Caption = "Production In Out Detail (" & CmbGroup.Text & ")"
        End If
      Else
        If RdoIn.Value = True Then
            RptReportViewer.Caption = "Production In Summary (" & CmbGroup.Text & ")"
        ElseIf RdoOut.Value = True Then
            RptReportViewer.Caption = "Production Out Summary (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Caption = "Production In Out Summary (" & CmbGroup.Text & ")"
        End If
      End If
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
        Case TxtFromStoreID.Name: If FunSelectFromStore(ssFunctionKey, True) = True Then TxtToStoreID.SetFocus
        Case TxtToStoreID.Name: If FunSelectToStore(ssFunctionKey, True) = True Then TxtCompanyID.SetFocus
        Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then TxtGroupID.SetFocus
        Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then TxtSubGroupID.SetFocus
        Case TxtSubGroupID.Name: If FunSelectSubGroup(ssFunctionKey, True) = True Then TxtProductID.SetFocus
        Case TxtProductID.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtUserNo.SetFocus
        Case TxtUserNo.Name: If FunSelectUser(ssFunctionKey, True) = True Then RdoIn.SetFocus
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
   SetWindowText Me.hWnd, "Production Register"
   CmbGroup.AddItem ("From Store Wise")
   CmbGroup.AddItem ("To Store Wise")
   CmbGroup.AddItem ("Company Wise")
   CmbGroup.AddItem ("Group Wise")
   CmbGroup.AddItem ("SubGroup Wise")
   CmbGroup.AddItem ("Product Wise")
   CmbGroup.AddItem ("Date Wise")
   CmbGroup.AddItem ("Invoice Wise")
   CmbGroup.AddItem ("User Wise")
   CmbGroup.ListIndex = 0
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
   'Set RsReport = Nothing
   Set RptProductionRegister = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Function SetReport() As Boolean
   On Error GoTo ErrorHandler
   SetReport = False
   Me.MousePointer = vbHourglass
   Dim RsReport As New ADODB.Recordset
   
   If RdoDetail.Value = True Then
      Select Case CmbGroup.Text
         Case "From Store Wise"
            Set RptReportViewer.Report = New CrpProductionDetailFromStoreWise
         Case "To Store Wise"
            Set RptReportViewer.Report = New CrpProductionDetailToStoreWise
         Case "Company Wise"
            Set RptReportViewer.Report = New CrpProductionDetailCompanyWise
         Case "Group Wise"
            Set RptReportViewer.Report = New CrpProductionDetailGroupWise
         Case "SubGroup Wise"
            Set RptReportViewer.Report = New CrpProductionDetailSubGroupWise
         Case "Product Wise"
            Set RptReportViewer.Report = New CrpProductionDetailProductWise
         Case "Date Wise"
            Set RptReportViewer.Report = New CrpProductionDetailDateWise
         Case "Invoice Wise"
            Set RptReportViewer.Report = New CrpProductionDetailInvoiceWise
         Case "User Wise"
            Set RptReportViewer.Report = New CrpProductionDetailUserWise
        End Select
   Else
      Select Case CmbGroup.Text
         Case "From Store Wise"
            Set RptReportViewer.Report = New CrpProductionSummaryFromStoreWise
         Case "To Store Wise"
            Set RptReportViewer.Report = New CrpProductionSummaryToStoreWise
         Case "Company Wise"
            Set RptReportViewer.Report = New CrpProductionSummaryCompanyWise
         Case "Group Wise"
            Set RptReportViewer.Report = New CrpProductionSummaryGroupWise
         Case "SubGroup Wise"
            Set RptReportViewer.Report = New CrpProductionSummarySubGroupWise
         Case "Product Wise"
            Set RptReportViewer.Report = New CrpProductionSummaryProductWise
         Case "Date Wise"
            Set RptReportViewer.Report = New CrpProductionSummaryDateWise
         Case "Invoice Wise"
            Set RptReportViewer.Report = New CrpProductionSummaryInvoiceWise
         Case "User Wise"
            Set RptReportViewer.Report = New CrpProductionSummaryUserWise
        End Select
   End If
    
   Set RsReport = CN.Execute("EXEC ProdRptProductionRegister '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtFromStoreID.Text) = "", "Null", "'" & TxtFromStoreID.Text & "'") & "," & IIf(Trim(TxtToStoreID.Text) = "", "Null", "'" & TxtToStoreID.Text & "'") & "," & IIf(Trim(TxtUserNo.Text) = "", "Null", "'" & TxtUserNo.Text & "'") & vbCrLf _
                                                                      & "," & IIf(Trim(TxtCompanyID.Text) = "", "Null", "'" & TxtCompanyID.Text & "'") & "," & IIf(Trim(TxtGroupID.Text) = "", "Null", "'" & TxtGroupID.Text & "'") & "," & IIf(Trim(TxtSubGroupID.Text) = "", "Null", "'" & TxtSubGroupID.Text & "'") & "," & IIf(Trim(TxtProductID.Text) = "", "Null", "'" & TxtProductID.Text & "'") & "," & IIf(RdoBoth.Value = True, "Null", IIf(RdoIn.Value = True, 0, 1)))
   If RsReport.BOF Then
      MsgBox "No record exists.", vbInformation, Me.Caption
      Me.MousePointer = vbDefault
      Exit Function
   End If
   
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport
    
   If RdoDetail.Value = True Then
        If RdoIn.Value = True Then
            RptReportViewer.Report.ReportTitle = "Production In Detail (" & CmbGroup.Text & ")"
        ElseIf RdoOut.Value = True Then
            RptReportViewer.Report.ReportTitle = "Production Out Detail (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Report.ReportTitle = "Production In Out Detail (" & CmbGroup.Text & ")"
        End If
   Else
        If RdoIn.Value = True Then
            RptReportViewer.Report.ReportTitle = "Production In Summary (" & CmbGroup.Text & ")"
        ElseIf RdoOut.Value = True Then
            RptReportViewer.Report.ReportTitle = "Production Out Summary (" & CmbGroup.Text & ")"
        Else
            RptReportViewer.Report.ReportTitle = "Production In Out Summary (" & CmbGroup.Text & ")"
        End If
   End If
    
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue " Date From :" & Format(DtpFrom.DateValue, "dd/MM/yyyy") & " To : " & Format(DtpTo.DateValue, "dd/MM/yyyy")
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue ObjRegistry.DevelopedBy
'    RptReportViewer.Report.SelectPrinter objRegistry.DriverName, objRegistry.DeviceName, objRegistry.Port
'    RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
    'RptReportViewer.Report.PaperOrientation = crLandscape
   SetReport = True
   Me.MousePointer = vbDefault
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

Private Sub TxtFromStoreID_Change()
   If TxtFromStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtFromStoreID.Name Then Exit Sub
   If TxtFromStoreName.Text <> "" Then TxtFromStoreName.Text = ""
End Sub

Private Sub TxtFromStoreID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtFromStoreID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtFromStoreID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectFromStore(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectFromStore(ssButton, False)
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

Private Sub TxtProductID_Change()
   If ActiveControl.Name <> TxtProductID.Name Then Exit Sub
   If TxtProductName.Text <> "" Then TxtProductName.Text = ""
End Sub

Private Sub TxtProductID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtProductID.Text) = "" Then Exit Sub
   vTemp = Not FunSelectProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectProduct(ssValidate, False)
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

Private Sub TxtToStoreID_Change()
   If TxtToStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtToStoreID.Name Then Exit Sub
   If TxtToStoreName.Text <> "" Then TxtToStoreName.Text = ""
End Sub

Private Sub TxtToStoreID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtToStoreID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtToStoreID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectToStore(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectToStore(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtUserNo_Change()
   If TxtUserNo.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtUserNo.Name Then Exit Sub
   If TxtUserName.Text <> "" Then TxtUserName.Text = ""
End Sub

Private Sub TxtUserNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtUserNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtUserNo.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectUser(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectUser(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

