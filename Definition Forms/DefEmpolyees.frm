VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form DefEmpolyees 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "DefEmpolyees.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkHideLockEmployee 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Hide Lock Employee"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3105
      TabIndex        =   86
      Tag             =   "NC"
      Top             =   1620
      Width           =   1860
   End
   Begin VB.TextBox TxtFinePerMinute 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   13590
      MaxLength       =   4
      TabIndex        =   84
      Top             =   8850
      Width           =   855
   End
   Begin VB.TextBox TxtFatherName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Height          =   330
      Left            =   7965
      TabIndex        =   5
      Top             =   4455
      Width           =   3780
   End
   Begin VB.TextBox TxtLoanInstallment 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   11745
      MaxLength       =   6
      TabIndex        =   16
      Top             =   7110
      Width           =   855
   End
   Begin VB.TextBox TxtStoreID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9360
      MaxLength       =   10
      TabIndex        =   1
      Top             =   2025
      Width           =   810
   End
   Begin VB.CheckBox ChkLockEmployee 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Lock Employee"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7980
      TabIndex        =   22
      Top             =   9495
      Width           =   1410
   End
   Begin VB.TextBox TxtHoursPerDay 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   11340
      MaxLength       =   2
      TabIndex        =   23
      Top             =   9360
      Width           =   855
   End
   Begin VB.TextBox TxtCreditLimit 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   11835
      MaxLength       =   6
      TabIndex        =   13
      Top             =   6570
      Width           =   855
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
      Height          =   2850
      Left            =   12840
      TabIndex        =   53
      Top             =   1200
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
         Height          =   2445
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   54
         Tag             =   "NC"
         Text            =   "DefEmpolyees.frx":0ECA
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
         TabIndex        =   55
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.TextBox TxtCommission 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   11115
      MaxLength       =   5
      TabIndex        =   12
      Top             =   6570
      Width           =   720
   End
   Begin VB.TextBox TxtSalary 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10125
      MaxLength       =   6
      TabIndex        =   11
      Top             =   6570
      Width           =   990
   End
   Begin VB.TextBox TxtReference2 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9810
      MaxLength       =   20
      TabIndex        =   15
      Top             =   7110
      Width           =   1935
   End
   Begin VB.TextBox TxtReference1 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7965
      MaxLength       =   20
      TabIndex        =   14
      Top             =   7110
      Width           =   1845
   End
   Begin VB.TextBox TxtDepartmentName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   9285
      TabIndex        =   45
      Top             =   2685
      Width           =   2475
   End
   Begin VB.TextBox TxtDepartmentID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7965
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2685
      Width           =   930
   End
   Begin VB.TextBox TxtDesignationName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   9285
      TabIndex        =   41
      Top             =   3300
      Width           =   2475
   End
   Begin VB.TextBox TxtDesignationID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7965
      MaxLength       =   10
      TabIndex        =   3
      Top             =   3300
      Width           =   930
   End
   Begin VB.TextBox TxtFilter 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2963
      MaxLength       =   30
      TabIndex        =   27
      Tag             =   "NC"
      Top             =   1995
      Width           =   3975
   End
   Begin VB.TextBox TxtPrefix 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Left            =   7965
      MaxLength       =   2
      TabIndex        =   36
      Tag             =   "NC"
      Top             =   2025
      Width           =   525
   End
   Begin VB.TextBox TxtCNIC 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7965
      MaxLength       =   20
      TabIndex        =   10
      Top             =   6570
      Width           =   2160
   End
   Begin VB.TextBox TxtMobileNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9840
      MaxLength       =   25
      TabIndex        =   9
      Top             =   6045
      Width           =   1980
   End
   Begin VB.TextBox TxtPhoneNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7965
      MaxLength       =   30
      TabIndex        =   8
      Top             =   6045
      Width           =   1875
   End
   Begin VB.TextBox TxtCity 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7965
      MaxLength       =   30
      TabIndex        =   7
      Top             =   5520
      Width           =   3870
   End
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7965
      MaxLength       =   100
      TabIndex        =   6
      Top             =   5010
      Width           =   3870
   End
   Begin VB.TextBox TxtID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8490
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2025
      Width           =   825
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   7005
      Left            =   2085
      TabIndex        =   28
      Top             =   2325
      Width           =   4935
      ScrollBars      =   2
      _Version        =   196616
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
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
      stylesets(0).Picture=   "DefEmpolyees.frx":0F55
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   2
      Columns(0).Width=   1852
      Columns(0).Caption=   "Employee ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5794
      Columns(1).Caption=   "Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   8705
      _ExtentY        =   12356
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
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7965
      MaxLength       =   30
      TabIndex        =   4
      Top             =   3855
      Width           =   3855
   End
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   2265
      TabIndex        =   29
      Top             =   10035
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "New"
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
      MICON           =   "DefEmpolyees.frx":0F71
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   3585
      TabIndex        =   30
      Top             =   10035
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Change"
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
      MICON           =   "DefEmpolyees.frx":0F8D
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   4905
      TabIndex        =   31
      Top             =   10035
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Remove"
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
      MICON           =   "DefEmpolyees.frx":0FA9
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   8070
      TabIndex        =   24
      Top             =   10035
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
      MICON           =   "DefEmpolyees.frx":0FC5
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9390
      TabIndex        =   25
      Top             =   10035
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
      MICON           =   "DefEmpolyees.frx":0FE1
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10710
      TabIndex        =   26
      Top             =   10035
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
      MICON           =   "DefEmpolyees.frx":0FFD
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDesignation 
      Height          =   330
      Left            =   8925
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   3300
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
      MICON           =   "DefEmpolyees.frx":1019
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDepartment 
      Height          =   330
      Left            =   8925
      TabIndex        =   46
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   2685
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
      MICON           =   "DefEmpolyees.frx":1035
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   10530
      TabIndex        =   59
      Tag             =   "NC"
      Top             =   2025
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Height          =   315
      Left            =   10170
      TabIndex        =   60
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   2025
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "DefEmpolyees.frx":1051
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSalaryAccount 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8760
      TabIndex        =   62
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   7680
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
      MICON           =   "DefEmpolyees.frx":106D
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSalaryAccountNo 
      Height          =   315
      Left            =   7980
      TabIndex        =   17
      Top             =   7680
      Width           =   780
      _ExtentX        =   1376
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
   End
   Begin SITextBox.Txt TxtSalaryAccountName 
      Height          =   315
      Left            =   9120
      TabIndex        =   63
      Tag             =   "NC"
      Top             =   7680
      Width           =   1470
      _ExtentX        =   2593
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
   Begin JeweledBut.JeweledButton BtnSaleAccount 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   11370
      TabIndex        =   65
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   7680
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
      MICON           =   "DefEmpolyees.frx":1089
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSaleAccountNo 
      Height          =   315
      Left            =   10590
      TabIndex        =   18
      Top             =   7680
      Width           =   780
      _ExtentX        =   1376
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
   End
   Begin SITextBox.Txt TxtSaleAccountName 
      Height          =   315
      Left            =   11730
      TabIndex        =   66
      Tag             =   "NC"
      Top             =   7680
      Width           =   1470
      _ExtentX        =   2593
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
   Begin JeweledBut.JeweledButton BtnAdvanceAccount 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8760
      TabIndex        =   68
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   8265
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
      MICON           =   "DefEmpolyees.frx":10A5
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtAdvanceAccountNo 
      Height          =   315
      Left            =   7980
      TabIndex        =   19
      Top             =   8265
      Width           =   780
      _ExtentX        =   1376
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
   End
   Begin SITextBox.Txt TxtAdvanceAccountName 
      Height          =   315
      Left            =   9120
      TabIndex        =   69
      Tag             =   "NC"
      Top             =   8265
      Width           =   1470
      _ExtentX        =   2593
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
   Begin JeweledBut.JeweledButton BtnLoanAccount 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   11370
      TabIndex        =   71
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   8265
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
      MICON           =   "DefEmpolyees.frx":10C1
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtLoanAccountNo 
      Height          =   315
      Left            =   10590
      TabIndex        =   20
      Top             =   8265
      Width           =   780
      _ExtentX        =   1376
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
   End
   Begin SITextBox.Txt TxtLoanAccountName 
      Height          =   315
      Left            =   11730
      TabIndex        =   72
      Tag             =   "NC"
      Top             =   8265
      Width           =   1470
      _ExtentX        =   2593
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
   Begin JeweledBut.JeweledButton BtnAccruedAccount 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8760
      TabIndex        =   74
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   8850
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
      MICON           =   "DefEmpolyees.frx":10DD
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtAccruedAccountNo 
      Height          =   315
      Left            =   7980
      TabIndex        =   21
      Top             =   8850
      Width           =   780
      _ExtentX        =   1376
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
   End
   Begin SITextBox.Txt TxtAccruedAccountName 
      Height          =   315
      Left            =   9120
      TabIndex        =   75
      Tag             =   "NC"
      Top             =   8850
      Width           =   1470
      _ExtentX        =   2593
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
   Begin JeweledBut.JeweledButton BtnAddBimetricImage 
      CausesValidation=   0   'False
      Height          =   675
      Left            =   12210
      TabIndex        =   77
      TabStop         =   0   'False
      Tag             =   "B"
      Top             =   5535
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1191
      TX              =   "Add Bimetric Image"
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
      MICON           =   "DefEmpolyees.frx":10F9
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpOfficeTimeIn 
      Height          =   315
      Left            =   10680
      TabIndex        =   82
      Top             =   8850
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Format          =   309264386
      UpDown          =   -1  'True
      CurrentDate     =   39805.4166666667
   End
   Begin MSComCtl2.DTPicker DtpOfficeTimeOut 
      Height          =   315
      Left            =   12030
      TabIndex        =   87
      Top             =   8850
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Format          =   309264386
      UpDown          =   -1  'True
      CurrentDate     =   39805.8333333333
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Office Time Out"
      Height          =   195
      Left            =   12015
      TabIndex        =   88
      Top             =   8640
      Width           =   1110
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fine Per Minute"
      Height          =   195
      Left            =   13590
      TabIndex        =   85
      Top             =   8640
      Width           =   1110
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Office Time In"
      Height          =   195
      Left            =   10665
      TabIndex        =   83
      Top             =   8640
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      Height          =   195
      Left            =   7965
      TabIndex        =   81
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   9360
      TabIndex        =   80
      Top             =   1800
      Width           =   585
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   10530
      TabIndex        =   79
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      Height          =   195
      Left            =   7965
      TabIndex        =   78
      Top             =   4230
      Width           =   915
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accrued A/C No"
      Height          =   195
      Left            =   7980
      TabIndex        =   76
      Top             =   8640
      Width           =   1185
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loan A/C No"
      Height          =   195
      Left            =   10590
      TabIndex        =   73
      Top             =   8055
      Width           =   945
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance A/C No"
      Height          =   195
      Left            =   7980
      TabIndex        =   70
      Top             =   8055
      Width           =   1230
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale A/C No"
      Height          =   195
      Left            =   10590
      TabIndex        =   67
      Top             =   7470
      Width           =   900
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary A/C No"
      Height          =   195
      Left            =   7980
      TabIndex        =   64
      Top             =   7470
      Width           =   1020
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Installment"
      Height          =   195
      Left            =   11580
      TabIndex        =   61
      Top             =   6885
      Width           =   1155
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Working Hours Per Day"
      Height          =   195
      Left            =   9525
      TabIndex        =   58
      Top             =   9480
      Width           =   1680
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Limit"
      Height          =   195
      Left            =   11805
      TabIndex        =   57
      Top             =   6390
      Width           =   765
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
      Left            =   11520
      TabIndex        =   56
      Top             =   585
      Width           =   435
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Commission (%)"
      Height          =   195
      Left            =   10680
      TabIndex        =   52
      Top             =   6390
      Width           =   1080
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      Height          =   195
      Left            =   10125
      TabIndex        =   51
      Top             =   6390
      Width           =   435
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference 2"
      Height          =   195
      Left            =   9810
      TabIndex        =   50
      Top             =   6915
      Width           =   885
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference 1"
      Height          =   195
      Left            =   7965
      TabIndex        =   49
      Top             =   6915
      Width           =   885
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      Height          =   195
      Left            =   9285
      TabIndex        =   48
      Top             =   2475
      Width           =   1290
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department ID"
      Height          =   195
      Left            =   7965
      TabIndex        =   47
      Top             =   2475
      Width           =   1035
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation Name"
      Height          =   195
      Left            =   9285
      TabIndex        =   44
      Top             =   3045
      Width           =   1305
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation ID"
      Height          =   195
      Left            =   7965
      TabIndex        =   43
      Top             =   3045
      Width           =   1050
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employees"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   40
      Top             =   270
      Width           =   1965
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   498
      X2              =   498
      Y1              =   131
      Y2              =   624
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   2408
      TabIndex        =   39
      Top             =   2055
      Width           =   510
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNIC No."
      Height          =   195
      Left            =   7965
      TabIndex        =   38
      Top             =   6375
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No"
      Height          =   195
      Left            =   9840
      TabIndex        =   37
      Top             =   5850
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No"
      Height          =   195
      Left            =   7965
      TabIndex        =   35
      Top             =   5850
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Index           =   0
      Left            =   7965
      TabIndex        =   34
      Top             =   5310
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   7980
      TabIndex        =   33
      Top             =   4815
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   7965
      TabIndex        =   32
      Top             =   3660
      Width           =   420
   End
End
Attribute VB_Name = "DefEmpolyees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Fl As Long, Chunks As Integer
Public Fragment As Integer

Public Templ As DPFPTemplate
Dim Rs As ADODB.Recordset
Dim vMode As FormMode
Dim ssql As String
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.


Private Function FunSelectStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectStore = False: Exit Function
        TxtStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Stores where islock = 0 and StoreID=" & Val(TxtStoreID.Text)
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

' Biimetric Feature
Public Sub SetTemplete(ByVal Template As Object)
   Set Templ = Template
End Sub

Public Function GetTemplate() As Object
 ' Template can be empty. If so, then returns Nothing.
 If Templ Is Nothing Then
 Else
  Set GetTemplate = Templ
 End If
End Function

Private Sub BtnAddBimetricImage_Click()
   FrmBiometricImage.Show
End Sub

Private Sub ChkHideLockEmployee_Click()
   If ActiveControl.Name <> ChkHideLockEmployee.Name Then Exit Sub
   Call TxtFilter_Change
End Sub

Private Sub DtpOfficeTimeIn_Change()
   If DtpOfficeTimeIn.Visible = False Then Exit Sub
   If Me.ActiveControl.Name <> DtpOfficeTimeIn.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub DtpOfficeTimeOut_Change()
   If DtpOfficeTimeOut.Visible = False Then Exit Sub
   If Me.ActiveControl.Name <> DtpOfficeTimeOut.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub TxtStoreID_Change()
   If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then TxtStoreName.Text = ""
End Sub

Private Sub TxtStoreID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
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

Private Function FunSelectEmpDepartment(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchEmpDepartment.Show vbModal, Me
        If SchEmpDepartment.ParaOutEmpDepartmentID = "" Then FunSelectEmpDepartment = False: Exit Function
        TxtDepartmentID.Text = SchEmpDepartment.ParaOutEmpDepartmentID
    End If
    '---------------------------
    vStrSQL = " Select * FROM EmpDepartments where EmpDepartmentID=" & Val(TxtDepartmentID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtDepartmentName.Text = !EmpDepartment
          FunSelectEmpDepartment = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectEmpDepartment = False
          .Close
          TxtDepartmentID.Text = ""
          TxtDepartmentName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectDesignation(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchDesignation.Show vbModal, Me
        If SchDesignation.ParaOutDesignationID = "" Then FunSelectDesignation = False: Exit Function
        TxtDesignationID.Text = SchDesignation.ParaOutDesignationID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Designations where DesignationID=" & Val(TxtDesignationID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtDesignationName.Text = !Designation
          FunSelectDesignation = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectDesignation = False
          .Close
          TxtDesignationID.Text = ""
          TxtDesignationName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      TxtDepartmentID.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub BtnDepartment_Click()
   If FunSelectEmpDepartment(ssButton, False) = True Then
     TxtDesignationID.SetFocus
   Else
      TxtDepartmentID.SetFocus
   End If
End Sub

Private Sub BtnDesignation_Click()
   If FunSelectDesignation(ssButton, False) = True Then
     TxtName.SetFocus
   Else
      TxtDesignationID.SetFocus
   End If
End Sub

Private Sub ChkLockEmployee_Click()
   If ActiveControl.Name <> ChkLockEmployee.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name = Grid.Name Then Call Grid_DblClick: Exit Sub
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
             If BtnSave.Enabled Then BtnSave_Click
             KeyCode = 0
         Case vbKeyW
             If BtnClear.Enabled Then BtnClear_Click
             KeyCode = 0
         Case vbKeyH
            FraHelp.ZOrder 0
            FraHelp.Visible = True
            KeyCode = 0
         Case vbKeyQ
             If BtnClose.Enabled Then BtnClose_Click
             KeyCode = 0
         Case vbKeyN
             If BtnNew.Enabled Then BtnNew_Click
             KeyCode = 0
         Case vbKeyO
             If BtnOpen.Enabled Then BtnOpen_Click
             KeyCode = 0
         Case vbKeyR
             If BtnDelete.Enabled Then BtnDelete_Click
             KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, True) = True Then TxtStoreID.SetFocus
         Case TxtDepartmentID.Name: If FunSelectEmpDepartment(ssFunctionKey, True) = True Then TxtDesignationID.SetFocus
         Case TxtDesignationID.Name: If FunSelectDesignation(ssFunctionKey, True) = True Then TxtName.SetFocus
         Case TxtSalaryAccountNo.Name: If FunSelectSalaryAccount(ssFunctionKey, True) = True Then If TxtSaleAccountNo.Enabled Then TxtSaleAccountNo.SetFocus
         Case TxtSaleAccountNo.Name: If FunSelectSalaryAccount(ssFunctionKey, True) = True Then If TxtAdvanceAccountNo.Enabled Then TxtAdvanceAccountNo.SetFocus
         Case TxtAdvanceAccountNo.Name: If FunSelectSalaryAccount(ssFunctionKey, True) = True Then If TxtLoanAccountNo.Enabled Then TxtLoanAccountNo.SetFocus
         Case TxtLoanAccountNo.Name: If FunSelectSalaryAccount(ssFunctionKey, True) = True Then If TxtAccruedAccountNo.Enabled Then TxtAccruedAccountNo.SetFocus
         Case TxtAccruedAccountNo.Name: If FunSelectAccruedAccount(ssFunctionKey, True) = True Then ChkLockEmployee.SetFocus
      End Select
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" And Me.ActiveControl.Tag = "" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
   If Grid.Rows > 0 And BtnOpen.Enabled Then BtnOpen_Click
End Sub

Private Sub Grid_GotFocus()
   On Error GoTo ErrorHandler
   Call Grid_RowColChange(0, 0)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKeyA To vbKeyZ, vbKey0 To vbKey9, Asc("a") To Asc("z")
      TxtFilter.Text = Chr(KeyAscii): TxtFilter.SelStart = Len(TxtFilter.Text): TxtFilter.SetFocus
   End Select
End Sub

Private Sub BtnClear_Click()
   FormStatus = SelectionMode
End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniEmployees", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
   Dim vtbl As String
   If Rs.RecordCount > 0 Then
      If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
         Dim vid As String
         vid = Rs!EmpID
         vtbl = Common.ChildDataExists("Employees", "EmpId='" & vid & "'", "") ' Common.ChildDataExists("ChartoFAccounts", "AccountNo='" & vID & "'", "Salesman")
         If vtbl <> "" Then
            MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
         Exit Sub
      End If
      cn.BeginTrans
      Call ActivityLog("Employees", eDelete, , , vid)
      Rs.Delete
      cn.Execute ("Delete From ChartOfAccounts Where AccountNo = '" & vid & "'")
      'CN.Execute ("Delete From users Where EmpID = '" & vid & "'")
      cn.CommitTrans
      If Rs.RecordCount = 0 Then FormStatus = NewMode: Exit Sub
      Rs.MoveNext
      Grid.MoveNext
      If Rs.EOF Then Rs.MoveLast
   End If
   Exit Sub
ErrorHandler:
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnNew_Click()
  FormStatus = NewMode
End Sub

Private Sub BtnOpen_Click()
  On Error GoTo ErrorHandler
  If Rs.RecordCount > 0 Then
    If Rs.BOF = False And Rs.EOF = False Then
      FormStatus = OpenMode
    End If
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  If FunValidation = False Then Exit Sub
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniEmployees", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
  If vIsNewRecord = False Then Call ActivityLog("Employees", eEdit, , , TxtPrefix.Text & TxtID.Text)
  Set Rs = New ADODB.Recordset
  ssql = " Select * FROM Employees where EmpID = '" & TxtPrefix.Text & TxtID.Text & "'"
  Rs.CursorLocation = adUseClient
  Rs.Open ssql, cn, adOpenStatic, adLockOptimistic
  cn.BeginTrans
  If vIsNewRecord Then
    cn.Execute ("Insert into chartofaccounts values (" & _
    "'" & TxtPrefix.Text & TxtID.Text & "',1,'" & Replace(TxtName.Text, "'", "''") & "','Employees',2,'" & Replace(TxtAddress.Text, "'", "''") & "','63',0,0,1," & ChkLockEmployee.Value & ",1,0,' ',0,0,'" & Date & "',0)")
    Rs.AddNew
    Rs!EmpID = TxtPrefix.Text & TxtID.Text
    Rs!isChanged = 0
    'CN.Execute ("Insert into users values (" & CN.Execute("Select isnull(max(userno),0) + 1 from users").Fields(0) & ",'" & TxtName.Text & "','',0,0,0,1,'" & Rs!EmpID & "')")
  Else
    Rs!isChanged = 1
    Rs!IsSync = 0
    Rs!modified_on = Now
    cn.Execute ("Update Chartofaccounts set Accountname='" & Replace(TxtName.Text, "'", "''") & "',Narration = '" & Replace(TxtAddress.Text, "'", "''") & "', IsSync = 0, isLocked = " & ChkLockEmployee.Value & ", isChanged = 1 Where AccountNo = '" & Rs!EmpID & "'")
    'CN.Execute ("Update users set username='" & TxtName.Text & "' Where EmpID = '" & Rs!EmpID & "'")
  End If
  If Not (Templ Is Nothing) Then
      Rs("BiometricPattern").AppendChunk Templ.Serialize
  End If
  Rs!StoreID = IIf(TxtStoreID.Text = "", Null, TxtStoreID.Text)
  Rs!DepartmentID = TxtDepartmentID.Text
  Rs!DesignationID = TxtDesignationID.Text
  Rs!empname = TxtName.Text
  Rs!EmpFather = IIf(TxtFatherName.Text = "", Null, TxtFatherName.Text)
  Rs!Address = TxtAddress.Text
  Rs!City = TxtCity.Text
  Rs!Phone = IIf(TxtPhoneNo.Text = "", Null, TxtPhoneNo.Text)
  Rs!Mobile = IIf(TxtMobileNo.Text = "", Null, TxtMobileNo.Text)
  Rs!CNIC = IIf(TxtCNIC.Text = "", Null, TxtCNIC.Text)
  Rs!Salary = Val(TxtSalary.Text)
  Rs!Commission = Val(TxtCommission.Text)
  Rs!CreditLimit = Val(TxtCreditLimit.Text)
  Rs!Reference1 = IIf(TxtReference1.Text = "", Null, TxtReference1.Text)
  Rs!Reference2 = IIf(TxtReference2.Text = "", Null, TxtReference2.Text)
  Rs!LoanInstallment = Val(TxtLoanInstallment.Text)
  Rs!OfficeTimein = Date & " " & Format(DtpOfficeTimeIn.Value, "hh:mm")
  Rs!OfficeTimeOut = Date & " " & Format(DtpOfficeTimeOut.Value, "hh:mm")
  Rs!FinePerMinute = IIf(Val(TxtFinePerMinute.Text) = 0, Null, Val(TxtFinePerMinute.Text))
  Rs!HoursPerDay = Val(TxtHoursPerDay.Text)
  Rs!SalaryAccountNo = IIf(TxtSalaryAccountNo.Text = "", Null, TxtSalaryAccountNo.Text)
  Rs!SaleAccountNo = IIf(TxtSaleAccountNo.Text = "", Null, TxtSaleAccountNo.Text)
  Rs!AdvanceAccountNo = IIf(TxtAdvanceAccountNo.Text = "", Null, TxtAdvanceAccountNo.Text)
  Rs!LoanAccountNo = IIf(TxtLoanAccountNo.Text = "", Null, TxtLoanAccountNo.Text)
  Rs!AccruedAccountNo = IIf(TxtAccruedAccountNo.Text = "", Null, TxtAccruedAccountNo.Text)
  Rs!IsLockEmployee = ChkLockEmployee.Value
  
  Rs.Update
  If vIsNewRecord = True Then Call ActivityLog("Employees", eAdd, , , TxtPrefix.Text & TxtID.Text)
  cn.CommitTrans
  
  Set Rs = New ADODB.Recordset
  ssql = " Select E.* FROM Employees e inner join ChartOfAccounts C on C.AccountNo = e.empid where EmpName like '%" & Replace(TxtFilter.Text, "'", "''") & "%'" & IIf(ChkHideLockEmployee.Value = 1, " and islocked = 0 and isLockEmployee = 0 ", "") & " Order by EmpName"
  Rs.CursorLocation = adUseClient
  Rs.Open ssql, cn, adOpenDynamic, adLockOptimistic
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  If cn.Errors.Count > 0 Then cn.RollbackTrans
  Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
  If vIsNewRecord Then
    If Trim(TxtID.Text) = "" Then
      MsgBox "Please specify a Employee ID", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
    If Trim(TxtDepartmentID.Text) = "" Then
      MsgBox "Please specify a Department ID", vbExclamation, "Alert"
      If TxtDepartmentID.Enabled And TxtDepartmentID.Visible Then TxtDepartmentID.SetFocus
      Exit Function
    End If
    If Not IsNumeric(TxtID.Text) Then
      MsgBox "The Employee ID must be numeric", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
  End If
  If Trim(TxtName.Text) = "" Then
    MsgBox "Please specify a Employee name", vbExclamation, "Alert"
    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
    Exit Function
  End If
  If Trim(TxtAddress.Text) = "" Then
    MsgBox "Please specify a Address", vbExclamation, "Alert"
    If TxtAddress.Enabled And TxtAddress.Visible Then TxtAddress.SetFocus
    Exit Function
  End If
  If Trim(TxtCity.Text) = "" Then
    MsgBox "Please specify a City", vbExclamation, "Alert"
    If TxtCity.Enabled And TxtCity.Visible Then TxtCity.SetFocus
    Exit Function
  End If
  If Trim(TxtSalary.Text) = "" Then
    MsgBox "Please specify a Salary", vbExclamation, "Alert"
    If TxtSalary.Enabled And TxtSalary.Visible Then TxtSalary.SetFocus
    Exit Function
  End If
  If TxtID.Enabled = True And cn.Execute("select count(*) from chartofaccounts where accountno = '" & TxtPrefix.Text & TxtID.Text & "'").Fields(0) > 0 Then
    MsgBox "This ID already exists. A new ID has been generated. Please save again", vbExclamation, "Alert"
    TxtID.Text = FunGetMaxID
    TxtID.SetFocus
    Exit Function
  End If
  'All Ok, now validation is success
  FunValidation = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

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
   SetWindowText Me.hWnd, "Employees"
   HelpLocation Me
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.Open "Select * FROM Employees", cn, adOpenStatic, adLockOptimistic
   Grid.Columns("ID").DataField = "EmpID"
   Grid.Columns("Name").DataField = "EmpName"
   FormStatus = NewMode
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Property Get FormStatus() As FormMode
  'Nothing
  FormStatus = vMode
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
'Based upon the value of vNewValue, we shall decide what controls to enable/disable
  On Error GoTo ErrorHandler
  vMode = vNewValue
  Select Case vNewValue
    Case Is = NewMode
      Call SubClearFields(True)
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtPrefix.Enabled = False
      TxtPrefix.Text = "63"
      TxtID.Text = FunGetMaxID
      TxtID.Enabled = False
      TxtFilter.Text = ""
      Grid.Enabled = False
      Set Grid.DataSource = Rs
      TxtFilter.Enabled = False
      ChkHideLockEmployee.Enabled = False
      If TxtStoreID.Enabled And TxtDepartmentID.Visible Then TxtDepartmentID.SetFocus
      vIsNewRecord = True
    Case Is = OpenMode
      Call SubClearFields(True)
      Call Grid_RowColChange(0, 0)
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnClear.Enabled = True
      Grid.Enabled = False
      TxtID.Enabled = False
      TxtFilter.Enabled = False
      ChkHideLockEmployee.Enabled = False
      TxtDesignationID.SetFocus
      vIsNewRecord = False
    Case Is = ChangeMode
      BtnSave.Enabled = True
    Case Is = SelectionMode
      Grid.Enabled = True
      Call SubClearFields(False)
      Call Grid_RowColChange(0, 0)
      TxtFilter.Enabled = True
      ChkHideLockEmployee.Enabled = True
      BtnNew.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnSave.Enabled = False
      BtnClear.Enabled = False
      Grid.SetFocus
      'TxtFilter.Text = ""
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo ErrorHandler
 
  If Rs.RecordCount > 0 And Grid.Enabled Then
      TxtID.Text = Right(Grid.Columns("ID").Text, Len(Grid.Columns("ID").Text) - 2)
      TxtName.Text = Grid.Columns("Name").Text
      TxtAddress.Text = Rs!Address
      TxtCity.Text = Rs!City
      TxtPhoneNo.Text = IIf(IsNull(Rs!Phone), "", Rs!Phone)
      TxtFatherName.Text = IIf(IsNull(Rs!EmpFather), "", Rs!EmpFather)
      TxtMobileNo.Text = IIf(IsNull(Rs!Mobile), "", Rs!Mobile)
      TxtCNIC.Text = IIf(IsNull(Rs!CNIC), "", Rs!CNIC)
      TxtSalary.Text = Rs!Salary
      TxtCreditLimit.Text = Rs!CreditLimit
      TxtLoanInstallment.Text = Rs!LoanInstallment
      If (IsNull(Rs!OfficeTimein)) Then
         DtpOfficeTimeIn.Value = Time
      Else
         DtpOfficeTimeIn.Value = Rs!OfficeTimein
      End If
      
      If (IsNull(Rs!OfficeTimeOut)) Then
         DtpOfficeTimeOut.Value = Time
      Else
         DtpOfficeTimeOut.Value = Rs!OfficeTimeOut
      End If
      TxtFinePerMinute.Text = IIf(IsNull(Rs!FinePerMinute), "", Rs!FinePerMinute)
      TxtCommission.Text = IIf(IsNull(Rs!Commission), "", Rs!Commission)
      TxtReference1.Text = IIf(IsNull(Rs!Reference1), "", Rs!Reference1)
      TxtReference2.Text = IIf(IsNull(Rs!Reference2), "", Rs!Reference2)
      TxtHoursPerDay.Text = IIf(IsNull(Rs!HoursPerDay), "", Rs!HoursPerDay)
      TxtDesignationID.Text = Rs!DesignationID
      If Trim(TxtDesignationID.Text) <> "" Then
         TxtDesignationName.Text = cn.Execute("select dbo.FunGetDesignation(" & Val(TxtDesignationID.Text) & ")").Fields(0).Value
      Else
         TxtDesignationName.Text = ""
      End If
      TxtDepartmentID.Text = Rs!DepartmentID
      If Trim(TxtDepartmentID.Text) <> "" Then
         TxtDepartmentName.Text = cn.Execute("select dbo.FunGetEmpDepartment(" & Val(TxtDepartmentID.Text) & ")").Fields(0).Value
      Else
         TxtDepartmentName.Text = ""
      End If
      TxtStoreID.Text = IIf(IsNull(Rs!StoreID), "", Rs!StoreID)
      If Trim(TxtStoreID.Text) <> "" Then
         TxtStoreName.Text = cn.Execute("select StoreName from Stores where StoreID = '" & Val(TxtStoreID.Text) & "'").Fields(0).Value
      Else
         TxtStoreName.Text = ""
      End If
      TxtSalaryAccountNo.Text = IIf(IsNull(Rs!SalaryAccountNo), "", Rs!SalaryAccountNo)
      If Trim(TxtSalaryAccountNo.Text) <> "" Then
         TxtSalaryAccountName.Text = cn.Execute("Select AccountName from ChartofAccounts Where AccountNo = '" & TxtSalaryAccountNo.Text & "'").Fields(0)
      Else
         TxtSalaryAccountName.Text = ""
      End If
      TxtSaleAccountNo.Text = IIf(IsNull(Rs!SaleAccountNo), "", Rs!SaleAccountNo)
      If Trim(TxtSaleAccountNo.Text) <> "" Then
         TxtSaleAccountName.Text = cn.Execute("Select AccountName from ChartofAccounts Where AccountNo = '" & TxtSaleAccountNo.Text & "'").Fields(0)
      Else
         TxtSaleAccountName.Text = ""
      End If
      TxtAdvanceAccountNo.Text = IIf(IsNull(Rs!AdvanceAccountNo), "", Rs!AdvanceAccountNo)
      If Trim(TxtAdvanceAccountNo.Text) <> "" Then
         TxtAdvanceAccountName.Text = cn.Execute("Select AccountName from ChartofAccounts Where AccountNo = '" & TxtAdvanceAccountNo.Text & "'").Fields(0)
      Else
         TxtAdvanceAccountName.Text = ""
      End If
      TxtLoanAccountNo.Text = IIf(IsNull(Rs!LoanAccountNo), "", Rs!LoanAccountNo)
      If Trim(TxtLoanAccountNo.Text) <> "" Then
         TxtLoanAccountName.Text = cn.Execute("Select AccountName from ChartofAccounts Where AccountNo = '" & TxtLoanAccountNo.Text & "'").Fields(0)
      Else
         TxtLoanAccountName.Text = ""
      End If
      TxtAccruedAccountNo.Text = IIf(IsNull(Rs!AccruedAccountNo), "", Rs!AccruedAccountNo)
      If Trim(TxtAccruedAccountNo.Text) <> "" Then
         TxtAccruedAccountName.Text = cn.Execute("Select AccountName from ChartofAccounts Where AccountNo = '" & TxtAccruedAccountNo.Text & "'").Fields(0)
      Else
         TxtAccruedAccountName.Text = ""
      End If
      ChkLockEmployee.Value = Abs(Rs!IsLockEmployee)
      If IsNull(Rs!BiometricPattern) Then
         Set Templ = Nothing
      Else
         With cn.Execute("Select BiometricPattern from Employees where EmpID = '" & Rs!EmpID & "'")
            If .RecordCount > 0 Then
               'ReDim blob(0)
               'ReDim blob(Rs("BiometricPattern").ActualSize)
                'put the picture data from the database in the array
               'blob() = Rs("BiometricPattern").GetChunk(Rs("BiometricPattern").ActualSize)
               ' Template can be empty, it must be created first.
               Set Templ = New DPFPTemplate
               ' Import binary data to template.
               Templ.Deserialize .Fields(0).GetChunk(.Fields(0).ActualSize)
            End If
            .Close
         End With
      End If
      
'      Open App.Path & "\abc.fpt" For Binary As #1
'      ReDim blob(LOF(1))
'      Get #1, , blob()
'      Close #1
'      ' Template can be empty, it must be created first.
'      Set Templ = New DPFPTemplate
'      ' Import binary data to template.
'      Templ.Deserialize blob

   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearFields(Enable As Boolean)
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is TextBox Then
         If ctl.Tag = "" Then
            ctl.Text = ""
            ctl.Enabled = Enable
         End If
      ElseIf TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then
            ctl.Text = ""
            ctl.Enabled = Enable
         End If
      ElseIf TypeOf ctl Is OptionButton Then
         If ctl.Tag = "" Then ctl.Enabled = Enable
      ElseIf TypeOf ctl Is CheckBox Then
         If ctl.Tag = "" Then ctl.Enabled = Enable
      ElseIf TypeOf ctl Is JeweledButton Then
         If ctl.Tag = "B" Then ctl.Enabled = Enable
      End If
   Next
   Set Templ = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtFilter_Change()
   On Error GoTo ErrorHandler
   'If ActiveControl.Name <> TxtFilter.Name Then Exit Sub
   'If Trim(TxtFilter.Text) = "" Then Grid.MoveFirst: Exit Sub
   'Rs.Find "Empname like '" & Replace(TxtFilter.Text, "'", "''") & "%'", , adSearchForward, 1
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
'   Rs.Open " Select * FROM Employees where EmpName like '%" & Replace(TxtFilter.Text, "'", "''") & "%' Order by EmpName", cn, adOpenStatic, adLockOptimistic
   ssql = " Select E.* FROM Employees e inner join ChartOfAccounts C on C.AccountNo = e.empid where EmpName like '%" & Replace(TxtFilter.Text, "'", "''") & "%'" & IIf(ChkHideLockEmployee.Value = 1, " and islocked = 0 and isLockEmployee = 0 ", "") & " Order by EmpName"
   Rs.Open ssql, cn, adOpenStatic, adLockOptimistic
   Set Grid.DataSource = Rs
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As String
   FunGetMaxID = cn.Execute("Select isnull(max(cast(substring(cast(accountno as varchar(5)),3,10) as smallint)),0) + 1 from chartofaccounts Where AccountNo like '63%' and isdetailed=1").Fields(0)
End Function

Private Sub TxtMobileNo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtCNIC_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("-"), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtName_Change()
   If Me.ActiveControl.Name <> TxtName.Name Then Exit Sub
   TxtFilter.Text = TxtName.Text
End Sub

Private Sub TxtPhoneNo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtDepartmentID_Change()
   If ActiveControl.Name <> TxtDepartmentID.Name Then Exit Sub
   If TxtDepartmentName.Text <> "" Then TxtDepartmentName.Text = ""
End Sub

Private Sub TxtDepartmentID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtDepartmentID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtDepartmentName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectEmpDepartment(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectEmpDepartment(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDesignationID_Change()
   If ActiveControl.Name <> TxtDesignationID.Name Then Exit Sub
   If TxtDesignationName.Text <> "" Then TxtDesignationName.Text = ""
End Sub

Private Sub TxtDesignationID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtDesignationID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtDesignationName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectDesignation(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectDesignation(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSalaryAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInWhereClause = ""
        SchAccounts.ParaInDetail = ""
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectSalaryAccount = False: Exit Function
        TxtSalaryAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select * FROM ChartOfAccounts where AccountNo='" & TxtSalaryAccountNo.Text & "'"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSalaryAccountName.Text = !AccountName
          FunSelectSalaryAccount = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSalaryAccount = False
          .Close
          TxtSalaryAccountNo.Text = ""
          TxtSalaryAccountName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtSalaryAccountNo_Change()
   If TxtSalaryAccountNo.Visible = False Then Exit Sub
   If TxtSalaryAccountName.Text <> "" Then TxtSalaryAccountName.Text = ""
End Sub

Private Sub TxtSalaryAccountNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSalaryAccountNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSalaryAccountNo.Text = "" Then Exit Sub
   If TxtSalaryAccountName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSalaryAccount(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSalaryAccount(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSalaryAccount_Click()
   If FunSelectSalaryAccount(ssButton, False) = True Then
      If TxtSaleAccountNo.Enabled Then TxtSaleAccountNo.SetFocus
   Else
      If TxtSalaryAccountNo.Enabled Then TxtSalaryAccountNo.SetFocus
   End If
End Sub

Private Function FunSelectSaleAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = ""
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectSaleAccount = False: Exit Function
        TxtSaleAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select * FROM ChartOfAccounts where AccountNo = '" & TxtSaleAccountNo.Text & "'"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSaleAccountName.Text = !AccountName
          FunSelectSaleAccount = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSaleAccount = False
          .Close
          TxtSaleAccountNo.Text = ""
          TxtSaleAccountName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtSaleAccountNo_Change()
   If TxtSaleAccountNo.Visible = False Then Exit Sub
   If TxtSaleAccountName.Text <> "" Then TxtSaleAccountName.Text = ""
End Sub

Private Sub TxtSaleAccountNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSaleAccountNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSaleAccountNo.Text = "" Then Exit Sub
   If TxtSaleAccountName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSaleAccount(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSaleAccount(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSaleAccount_Click()
   If FunSelectSaleAccount(ssButton, False) = True Then
      If TxtAdvanceAccountNo.Enabled Then TxtAdvanceAccountNo.SetFocus
   Else
      If TxtSaleAccountNo.Enabled Then TxtSaleAccountNo.SetFocus
   End If
End Sub

Private Function FunSelectAdvanceAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = ""
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectAdvanceAccount = False: Exit Function
        TxtAdvanceAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select * FROM ChartOfAccounts where AccountNo = '" & TxtAdvanceAccountNo.Text & "'"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtAdvanceAccountName.Text = !AccountName
          FunSelectAdvanceAccount = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectAdvanceAccount = False
          .Close
          TxtAdvanceAccountNo.Text = ""
          TxtAdvanceAccountName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtAdvanceAccountNo_Change()
   If TxtAdvanceAccountNo.Visible = False Then Exit Sub
   If TxtAdvanceAccountName.Text <> "" Then TxtAdvanceAccountName.Text = ""
End Sub

Private Sub TxtAdvanceAccountNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtAdvanceAccountNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtAdvanceAccountNo.Text = "" Then Exit Sub
   If TxtAdvanceAccountName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectAdvanceAccount(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectAdvanceAccount(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAdvanceAccount_Click()
   If FunSelectAdvanceAccount(ssButton, False) = True Then
      If TxtLoanAccountNo.Enabled Then TxtLoanAccountNo.SetFocus
   Else
      If TxtAdvanceAccountNo.Enabled Then TxtAdvanceAccountNo.SetFocus
   End If
End Sub

Private Function FunSelectLoanAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInDetail = ""
        SchAccounts.ParaInWhereClause = ""
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectLoanAccount = False: Exit Function
        TxtLoanAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select * FROM ChartOfAccounts where AccountNo = '" & TxtLoanAccountNo.Text & "'"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtLoanAccountName.Text = !AccountName
          FunSelectLoanAccount = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectLoanAccount = False
          .Close
          TxtLoanAccountNo.Text = ""
          TxtLoanAccountName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtLoanAccountNo_Change()
   If TxtLoanAccountNo.Visible = False Then Exit Sub
   If TxtLoanAccountName.Text <> "" Then TxtLoanAccountName.Text = ""
End Sub

Private Sub TxtLoanAccountNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtLoanAccountNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtLoanAccountNo.Text = "" Then Exit Sub
   If TxtLoanAccountName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectLoanAccount(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectLoanAccount(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnLoanAccount_Click()
   If FunSelectLoanAccount(ssButton, False) = True Then
      If TxtAccruedAccountNo.Enabled Then TxtAccruedAccountNo.SetFocus
   Else
      If TxtLoanAccountNo.Enabled Then TxtLoanAccountNo.SetFocus
   End If
End Sub

Private Function FunSelectAccruedAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInWhereClause = ""
        SchAccounts.ParaInDetail = ""
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectAccruedAccount = False: Exit Function
        TxtAccruedAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = "Select * FROM ChartOfAccounts where AccountNo = '" & TxtAccruedAccountNo.Text & "'"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtAccruedAccountName.Text = !AccountName
          FunSelectAccruedAccount = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectAccruedAccount = False
          .Close
          TxtAccruedAccountNo.Text = ""
          TxtAccruedAccountName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtAccruedAccountNo_Change()
   If TxtAccruedAccountNo.Visible = False Then Exit Sub
   If TxtAccruedAccountName.Text <> "" Then TxtAccruedAccountName.Text = ""
End Sub

Private Sub TxtAccruedAccountNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtAccruedAccountNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtAccruedAccountNo.Text = "" Then Exit Sub
   If TxtAccruedAccountName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectAccruedAccount(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectAccruedAccount(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnAccruedAccount_Click()
   If FunSelectAccruedAccount(ssButton, False) = True Then
      If BtnSave.Enabled Then ChkLockEmployee.SetFocus
   Else
      If TxtAccruedAccountNo.Enabled Then TxtAccruedAccountNo.SetFocus
   End If
End Sub

