VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptSalestaxRegister 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptSalestaxRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbTaxType 
      Height          =   315
      Left            =   11430
      Style           =   2  'Dropdown List
      TabIndex        =   60
      Top             =   4770
      Width           =   1950
   End
   Begin MSComCtl2.DTPicker DTPTimeFrom 
      Height          =   345
      Left            =   9000
      TabIndex        =   58
      Top             =   5895
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "hh:mm:ss"
      Format          =   137232386
      UpDown          =   -1  'True
      CurrentDate     =   36494
   End
   Begin VB.ComboBox CmbSortType 
      Height          =   315
      ItemData        =   "RptSalestaxRegister.frx":0ECA
      Left            =   10905
      List            =   "RptSalestaxRegister.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   6780
      Width           =   1275
   End
   Begin VB.ComboBox CmbSortName 
      Height          =   315
      ItemData        =   "RptSalestaxRegister.frx":0ECE
      Left            =   8880
      List            =   "RptSalestaxRegister.frx":0ED0
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   6780
      Width           =   1815
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      ItemData        =   "RptSalestaxRegister.frx":0ED2
      Left            =   8055
      List            =   "RptSalestaxRegister.frx":0ED4
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4410
      Width           =   3390
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   11490
      TabIndex        =   32
      Top             =   4410
      Width           =   1860
      Begin VB.OptionButton RdoSummary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Summary"
         Height          =   255
         Left            =   900
         TabIndex        =   13
         Top             =   10
         Width           =   960
      End
      Begin VB.OptionButton RdoDetail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detail"
         Height          =   255
         Left            =   75
         TabIndex        =   12
         Top             =   10
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   8085
      TabIndex        =   31
      Top             =   4035
      Width           =   5265
      Begin VB.OptionButton RdoNet 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Net Sale"
         Height          =   300
         Left            =   4125
         TabIndex        =   10
         Top             =   15
         Width           =   915
      End
      Begin VB.OptionButton RdoInv 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sale"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   15
         Width           =   630
      End
      Begin VB.OptionButton RdoReturn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sale Return"
         Height          =   255
         Left            =   1215
         TabIndex        =   8
         Top             =   15
         Width           =   1140
      End
      Begin VB.OptionButton RdoBoth 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Both"
         Height          =   255
         Left            =   2700
         TabIndex        =   9
         Top             =   15
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   11625
      TabIndex        =   18
      Top             =   7680
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
      MICON           =   "RptSalestaxRegister.frx":0ED6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8865
      TabIndex        =   16
      Top             =   7650
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
      MICON           =   "RptSalestaxRegister.frx":0EF2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10230
      TabIndex        =   17
      Top             =   7680
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
      MICON           =   "RptSalestaxRegister.frx":0F0E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   5790
      TabIndex        =   33
      Top             =   420
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   1605
      TabIndex        =   6
      Top             =   7403
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   20
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
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   2625
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7403
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
      MICON           =   "RptSalestaxRegister.frx":0F2A
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   2985
      TabIndex        =   19
      Top             =   7403
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   556
      Appearance      =   0
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
      Height          =   345
      Left            =   9000
      TabIndex        =   14
      Top             =   5490
      Width           =   1305
      _Version        =   65543
      _ExtentX        =   2302
      _ExtentY        =   609
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
      Left            =   10890
      TabIndex        =   15
      Top             =   5490
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
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2625
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5513
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
      MICON           =   "RptSalestaxRegister.frx":0F46
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   1605
      TabIndex        =   3
      Top             =   5513
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   2985
      TabIndex        =   27
      Tag             =   "nc"
      Top             =   5513
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
      Left            =   2625
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   6143
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
      MICON           =   "RptSalestaxRegister.frx":0F62
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubGroupID 
      Height          =   315
      Left            =   1605
      TabIndex        =   4
      Top             =   6143
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtSubGroupName 
      Height          =   315
      Left            =   2985
      TabIndex        =   29
      Tag             =   "nc"
      Top             =   6143
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
      Left            =   2625
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4883
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
      MICON           =   "RptSalestaxRegister.frx":0F7E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   1605
      TabIndex        =   2
      Top             =   4883
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtCompanyName 
      Height          =   315
      Left            =   2985
      TabIndex        =   25
      Tag             =   "nc"
      Top             =   4883
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
      Left            =   2625
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4253
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
      MICON           =   "RptSalestaxRegister.frx":0F9A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtUserNo 
      Height          =   315
      Left            =   1605
      TabIndex        =   1
      Top             =   4253
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
   Begin SITextBox.Txt TxtUserName 
      Height          =   315
      Left            =   2985
      TabIndex        =   23
      Tag             =   "nc"
      Top             =   4253
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
   Begin JeweledBut.JeweledButton BtnCustomer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2625
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3623
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
      MICON           =   "RptSalestaxRegister.frx":0FB6
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPartyID 
      Height          =   315
      Left            =   1605
      TabIndex        =   0
      Top             =   3623
      Width           =   1020
      _ExtentX        =   1799
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
      IntegralPoint   =   10
   End
   Begin SITextBox.Txt TxtPartyName 
      Height          =   315
      Left            =   2985
      TabIndex        =   21
      Tag             =   "nc"
      Top             =   3623
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
   Begin JeweledBut.JeweledButton BtnBrand 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2625
      TabIndex        =   54
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   6773
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
      MICON           =   "RptSalestaxRegister.frx":0FD2
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBrandID 
      Height          =   315
      Left            =   1605
      TabIndex        =   5
      Top             =   6773
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
      Left            =   2985
      TabIndex        =   55
      Tag             =   "nc"
      Top             =   6773
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
   Begin MSComCtl2.DTPicker DTPTimeTo 
      Height          =   315
      Left            =   10890
      TabIndex        =   59
      Top             =   5850
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "hh:mm:ss"
      Format          =   137428994
      UpDown          =   -1  'True
      CurrentDate     =   0.499988425925926
   End
   Begin VB.Label LblType 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Type"
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
      Left            =   10575
      TabIndex        =   61
      Top             =   4815
      Width           =   810
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
      Left            =   1605
      TabIndex        =   57
      Top             =   6578
      Width           =   765
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
      Left            =   2985
      TabIndex        =   56
      Top             =   6578
      Width           =   1050
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Type"
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
      Left            =   10905
      TabIndex        =   52
      Top             =   6555
      Width           =   840
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Name"
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
      Left            =   8880
      TabIndex        =   51
      Top             =   6555
      Width           =   900
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   899
      X2              =   899
      Y1              =   257
      Y2              =   427
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   531
      X2              =   531
      Y1              =   261
      Y2              =   429
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   899
      X2              =   533
      Y1              =   429
      Y2              =   429
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   531
      X2              =   899
      Y1              =   259
      Y2              =   259
   End
   Begin VB.Label Label14 
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
      Left            =   1605
      TabIndex        =   49
      Top             =   3398
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
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
      Left            =   2985
      TabIndex        =   48
      Top             =   3398
      Width           =   1335
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
      Left            =   2985
      TabIndex        =   47
      Top             =   4043
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
      Left            =   1605
      TabIndex        =   46
      Top             =   4043
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
      Left            =   1605
      TabIndex        =   45
      Top             =   5288
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
      Left            =   2985
      TabIndex        =   44
      Top             =   5288
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
      Left            =   1605
      TabIndex        =   43
      Top             =   5918
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
      Left            =   1605
      TabIndex        =   42
      Top             =   4673
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
      Left            =   2985
      TabIndex        =   41
      Top             =   4673
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
      Left            =   2985
      TabIndex        =   40
      Top             =   5918
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
      Left            =   1605
      TabIndex        =   39
      Top             =   7193
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
      Left            =   2985
      TabIndex        =   38
      Top             =   7193
      Width           =   1215
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
      Left            =   10890
      TabIndex        =   37
      Top             =   5310
      Width           =   705
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
      Left            =   9000
      TabIndex        =   36
      Top             =   5310
      Width           =   885
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SalesTax Register"
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
      TabIndex        =   35
      Top             =   270
      Width           =   2070
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   5790
      TabIndex        =   34
      Top             =   225
      Visible         =   0   'False
      Width           =   720
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
Attribute VB_Name = "RptSalestaxRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim sSql, vSelectedParameter As String
Dim Application1 As New CRAXDRT.Application


Private Function FunSelectCustomer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
'        SchAccounts.CmbFilter = "Customers"
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = " and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo Like '3%') and c.isLocked = 0"
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectCustomer = False: Exit Function
        TxtPartyID.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    VStrSQL = "Select c.AccountNo, c.AccountName as AccountName, Address, City, p.Description, isnull(p.isWholeSale,1) as isWholeSale" & vbCrLf _
         + " from ChartofAccounts c  " & vbCrLf _
         + " left outer join Parties p on p.partyid = c.AccountNo  " & vbCrLf _
         + " where c.AccountNo in  ( " & (TxtPartyID.Text) & " ) and (c.AccountNo like '6%' or c.AccountNo like '5%' or c.AccountNo Like '3%') and isDetailed = 1 and isLocked = 0"
    
    VStrSQL = VStrSQL + " union all Select EmpID, EmpName as AccountName, Address, City, '', 1 as isWholeSale" & vbCrLf _
         + " from Employees" & vbCrLf _
         + " where EmpID in ( " & (TxtPartyID.Text) & " ) and isLockEmployee = 0"
        
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtPartyName.Text = !AccountName
          FunSelectCustomer = True
          .Close
          Exit Function
      Else
          FunSelectCustomer = False
          .Close
          TxtPartyID.Text = ""
          TxtPartyName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtCode.Text) = "" Then Exit Function
    If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    VStrSQL = " SELECT p.productid, code, ProductName" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where p.productid = " & Val(TxtCode.Text) & " or code = '" & TxtCode.Text & "'"
  
   With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !ProductID
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
      RdoBoth.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnCustomer_Click()
   If FunSelectCustomer(ssButton, False) = True Then
      TxtUserNo.SetFocus
   Else
      TxtPartyID.SetFocus
   End If
End Sub

Private Sub BtnUser_Click()
   If FunSelectUser(ssButton, False) = True Then
      TxtCompanyID.SetFocus
   Else
      TxtUserNo.SetFocus
   End If
End Sub

Private Sub TxtProductName_Change()
   If ActiveControl.Name <> TxtProductName.Name Then Exit Sub
   If TxtProductID.Text <> "" Then TxtProductID.Text = ""
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
'      TxtCode.Text = ""
      TxtProductID.Text = ""
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

Private Sub BtnCompany_Click()
   If FunSelectCompany(ssButton, False) = True Then
      TxtGroupID.SetFocus
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

Private Sub TxtPartyID_Change()
   If TxtPartyID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtPartyID.Name Then Exit Sub
   If TxtPartyName.Text <> "" Then TxtPartyName.Text = ""
End Sub

Private Sub TxtPartyID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtPartyID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtPartyID.Text = "" Then Exit Sub
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

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
    If SetReport Then
      If RdoDetail.Value = True Then
          If RdoInv.Value = True Then
              RptReportViewer.Caption = "Salestax Detail (" & CmbGroup.Text & ")"
          ElseIf RdoReturn.Value = True Then
              RptReportViewer.Caption = "Salestax Return Detail (" & CmbGroup.Text & ")"
          Else
              RptReportViewer.Caption = "Salestax & Salestax Return Detail (" & CmbGroup.Text & ")"
          End If
      Else
          If RdoInv.Value = True Then
              RptReportViewer.Caption = "Salestax Summary (" & CmbGroup.Text & ")"
          ElseIf RdoReturn.Value = True Then
              RptReportViewer.Caption = "Salestax Return Summary (" & CmbGroup.Text & ")"
          Else
              RptReportViewer.Caption = "Salestax & Salestax Return Summary (" & CmbGroup.Text & ")"
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
        Case TxtPartyID.Name: If FunSelectCustomer(ssFunctionKey, True) = True Then TxtUserNo.SetFocus
        Case TxtUserNo.Name: If FunSelectUser(ssFunctionKey, True) = True Then TxtCompanyID.SetFocus
        Case TxtCompanyID.Name: If FunSelectCompany(ssFunctionKey, True) = True Then TxtGroupID.SetFocus
        Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then TxtSubGroupID.SetFocus
        Case TxtSubGroupID.Name: If FunSelectSubGroup(ssFunctionKey, True) = True Then TxtBrandID.SetFocus
        Case TxtBrandID.Name: If FunSelectBrand(ssFunctionKey, True) = True Then TxtCode.SetFocus
        Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then RdoInv.SetFocus
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
   SetWindowText Me.hWnd, "Salestax Register"
   
   CmbGroup.Clear
'   CmbGroup.AddItem ("Brand Wise")
'   CmbGroup.AddItem ("Company Wise")
'   CmbGroup.AddItem ("Customer Wise")
'   CmbGroup.AddItem ("Date Wise")
'   CmbGroup.AddItem ("Group Wise")
   CmbGroup.AddItem ("Invoice Wise")
   CmbGroup.AddItem ("Tax Type Wise")
'   CmbGroup.AddItem ("Product Wise")
'   CmbGroup.AddItem ("SubGroup Wise")
'   CmbGroup.AddItem ("User Wise")
   
   CmbSortName.Clear
   CmbSortName.AddItem "ProductID"
   CmbSortName.AddItem "ProductName"
   CmbSortType.Clear
   CmbSortType.AddItem "Ascending"
   CmbSortType.AddItem "Descending"
   CmbSortName.ListIndex = 0
   CmbSortType.ListIndex = 0
   
   
   CmbTaxType.Clear
   CmbTaxType.AddItem "All Types"
   CmbTaxType.AddItem "3rd Schedule"
   CmbTaxType.AddItem "Exempted"
   CmbTaxType.AddItem "Others"

   
   CmbGroup.ListIndex = 0
   CmbTaxType.ListIndex = 0
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
   Set RptSalestaxRegister = Nothing
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
        

   sSql = "EXEC ProdRptSalesTaxRegister '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'" & vbCrLf _
                              & "," & IIf(Trim(TxtPartyID.Text) = "", "'", "'" & TxtPartyID.Text) & "'," & IIf(Trim(TxtUserNo.Text) = "", "Null", "'" & TxtUserNo.Text & "'") & vbCrLf _
                              & "," & IIf(Trim(TxtCompanyID.Text) = "", "Null", "'" & TxtCompanyID.Text & "'") & "," & IIf(Trim(TxtGroupID.Text) = "", "Null", "'" & TxtGroupID.Text & "'") & "," & IIf(Trim(TxtSubGroupID.Text) = "", "Null", "'" & TxtSubGroupID.Text & "'") & "," & IIf(Trim(TxtBrandID.Text) = "", "Null", TxtBrandID.Text) & "," & IIf(Trim(TxtProductID.Text) = "", "Null", "'" & TxtProductID.Text & "'") & vbCrLf _
                              & "," & IIf(RdoBoth.Value = True Or RdoNet.Value = True, "Null", IIf(RdoInv.Value = True, 0, 1)) & "," & IIf(Trim(TxtProductID.Text) <> "", "Null", "'%" & TxtProductName.Text & "%'") & vbCrLf _
                              & "," & IIf(CmbTaxType.ListIndex = 1, 1, "null") & vbCrLf _
                              & "," & IIf(CmbTaxType.ListIndex = 2, 1, "null") & vbCrLf _
                              & "," & IIf(CmbTaxType.ListIndex = 3, 1, "null") & vbCrLf _
                              & "," & "'" & CmbSortName.Text & " " & CmbSortType.Text & "'" & vbCrLf _

'   Set RsReport = cn.Execute(sSql)
   
   If RdoDetail.Value = True Then
        Select Case CmbGroup.Text
            Case "Customer Wise"
                Set RptReportViewer.Report = New CrptSaleOrderDetailCustomerWise
            Case "Employee Wise"
                Set RptReportViewer.Report = New CrptSaleDetailEmpWise
            Case "User Wise"
                Set RptReportViewer.Report = New CrptSaleDetailUserWise
            Case "Company Wise"
                Set RptReportViewer.Report = New CrptSaleDetailCompanyWise
            Case "Group Wise"
                Set RptReportViewer.Report = New CrptSaleDetailGroupWise
            Case "SubGroup Wise"
                Set RptReportViewer.Report = New CrptSaleDetailSubGroupWise
            Case "Product Wise"
                Set RptReportViewer.Report = New CrptSaleDetailProductWise
            Case "Date Wise"
                Set RptReportViewer.Report = New CrptSaleDetailDateWise
            Case "Invoice Wise"
                Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\SaleReports\CrptSalestaxDetailInvoiceWise.rpt")
                'Set RptReportViewer.Report = New CrptSaleDetailInvoiceWise
            Case "Tax Type Wise"
                Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\SaleReports\CrptSalestaxDetailTaxTypeWise.rpt")
            Case "Invoice Wise Tax"
                Set RptReportViewer.Report = New CrptSaleDetailInvoiceWiseTax
            Case "Brand Wise"
                Set RptReportViewer.Report = New CrptSaleDetailBrandWise
        End Select
   Else
        Select Case CmbGroup.Text
            Case "Customer Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryCustomerWise
            Case "Employee Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryEmpWise
            Case "User Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryUserWise
            Case "Company Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryCompanyWise
            Case "Group Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryGroupWise
            Case "SubGroup Wise"
                Set RptReportViewer.Report = New CrptSaleSummarySubGroupWise
            Case "Product Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryProductNameWise
            Case "Date Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryDateWise
            Case "Invoice Wise"
                Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\SaleReports\CrptSalestaxSummaryInvoiceWise.rpt")
            Case "Tax Type Wise"
                Set RptReportViewer.Report = Application1.OpenReport(App.Path & "\reports\SaleReports\CrptSalestaxSummaryTaxTypeWise.rpt")
            Case "Brand Wise"
                Set RptReportViewer.Report = New CrptSaleSummaryBrandWise
        End Select
    End If
           
    Set RsReport = CN.Execute(sSql)
    
    If RsReport.BOF Then
        MsgBox "No record exists.", vbInformation, Me.Caption
        Me.MousePointer = vbDefault
        Exit Function
    End If
    RptReportViewer.Report.Database.SetDataSource RsReport
    
    If RdoDetail.Value = True Then
         RptReportViewer.Report.ReportTitle = "Salestax Detail (" & CmbGroup.Text & ")"
    Else
         RptReportViewer.Report.ReportTitle = "Salestax Summary (" & CmbGroup.Text & ")"
   End If
   
    
'    If RdoNet.Value Then
'      RptReportViewer.Report.DeleteGroup 1
'    End If
    Call SubSelectedParameterFields
    RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName & IIf(ObjRegistry.CompanyCity = "", "", " - " & ObjRegistry.CompanyCity)
    RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress)
    RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", "Phone # " & ObjRegistry.CompanyPhoneNo)
    RptReportViewer.Report.ParameterFields(5).AddCurrentValue "Date From " & Format(DtpFrom.DateValue, "dd/MM/yyyy") & " To " & Format(DtpTo.DateValue, "dd/MM/yyyy")
    RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
    RptReportViewer.Report.ParameterFields(6).AddCurrentValue vSelectedParameter
    RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
    
'    If ObjRegistry.IsPortrait = False Then RptReportViewer.Report.PaperOrientation = crLandscape
    
   RptReportViewer.Report.PaperOrientation = crPortrait
    
    'RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
    'RptReportViewer.Report.PaperOrientation = crLandscape
    SetReport = True
    Me.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
    Me.MousePointer = vbDefault
End Function

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

Private Function FunSelectUser(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchUser.Show vbModal, Me
        If SchUser.ParaOutUserNo = "" Then FunSelectUser = False: Exit Function
        TxtUserNo.Text = SchUser.ParaOutUserNo
    End If
    '---------------------------
    VStrSQL = " Select * FROM Users where UserNo=" & Val(TxtUserNo.Text)
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtUserName.Text = !UserName
          FunSelectUser = True
          .Close
          Exit Function
             FunSelectUser = True
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

Private Sub BtnBrand_Click()
   If FunSelectBrand(ssButton, False) = True Then
      TxtCode.SetFocus
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

Private Sub SubSelectedParameterFields()
   On Error GoTo ErrorHandler
   vSelectedParameter = ""
   If TxtUserName.Text <> "" Then vSelectedParameter = vSelectedParameter & "User: " & TxtUserNo.Text & " - " & TxtUserName.Text & " " & vbCrLf
   If TxtCompanyName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Company: " & TxtCompanyID.Text & " - " & TxtCompanyName.Text & " " & vbCrLf
   If TxtGroupName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Group: " & TxtGroupID.Text & " - " & TxtGroupName.Text & " " & vbCrLf
   If TxtSubGroupName.Text <> "" Then vSelectedParameter = vSelectedParameter & "SubGroup: " & TxtSubGroupID.Text & " - " & TxtSubGroupName.Text & " " & vbCrLf
   If TxtBrandName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Brand: " & TxtBrandID.Text & " - " & TxtBrandName.Text & " " & vbCrLf
   If TxtProductName.Text <> "" Then vSelectedParameter = vSelectedParameter & "Product: " & TxtProductID.Text & " - " & TxtProductName.Text & " " & vbCrLf
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

