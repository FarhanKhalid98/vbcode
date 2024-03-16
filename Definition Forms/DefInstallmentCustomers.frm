VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form DefInstallmentCustomers 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "DefInstallmentCustomers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNoOfInstallments 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9855
      MaxLength       =   2
      TabIndex        =   14
      Top             =   7590
      Width           =   2070
   End
   Begin VB.TextBox TxtPromiseDate 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7875
      MaxLength       =   2
      TabIndex        =   13
      Top             =   7590
      Width           =   1980
   End
   Begin VB.TextBox TxtReferFName2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   10980
      TabIndex        =   57
      Tag             =   "NC"
      Top             =   8970
      Width           =   2025
   End
   Begin VB.TextBox TxtReferID2 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7890
      MaxLength       =   10
      TabIndex        =   16
      Tag             =   "C"
      Top             =   8970
      Width           =   705
   End
   Begin VB.TextBox TxtReferName2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   8955
      TabIndex        =   56
      Tag             =   "NC"
      Top             =   8970
      Width           =   2025
   End
   Begin VB.TextBox TxtReferFName1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   10980
      TabIndex        =   51
      Tag             =   "NC"
      Top             =   8340
      Width           =   2025
   End
   Begin VB.TextBox TxtReferID1 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7890
      MaxLength       =   10
      TabIndex        =   15
      Tag             =   "C"
      Top             =   8340
      Width           =   705
   End
   Begin VB.TextBox TxtReferName1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   8955
      TabIndex        =   50
      Tag             =   "NC"
      Top             =   8340
      Width           =   2025
   End
   Begin VB.TextBox TxtPhone1 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7883
      MaxLength       =   30
      TabIndex        =   8
      Top             =   6225
      Width           =   2655
   End
   Begin VB.TextBox TxtPhone2 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10538
      MaxLength       =   30
      TabIndex        =   9
      Top             =   6225
      Width           =   2595
   End
   Begin VB.CheckBox ChkDateofBirth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11933
      TabIndex        =   47
      Top             =   6675
      Width           =   195
   End
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   7883
      MaxLength       =   100
      TabIndex        =   4
      Top             =   3765
      Width           =   5265
   End
   Begin VB.TextBox TxtCity 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7883
      MaxLength       =   30
      TabIndex        =   6
      Top             =   5655
      Width           =   2655
   End
   Begin VB.TextBox TxtMobileNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7883
      MaxLength       =   30
      TabIndex        =   10
      Top             =   6885
      Width           =   1980
   End
   Begin VB.TextBox TxtFName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7883
      MaxLength       =   30
      TabIndex        =   3
      Top             =   3180
      Width           =   5085
   End
   Begin VB.TextBox TxtAddress2 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   7883
      MaxLength       =   100
      TabIndex        =   5
      Top             =   4695
      Width           =   5265
   End
   Begin VB.TextBox TxtCast 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10538
      MaxLength       =   30
      TabIndex        =   7
      Top             =   5655
      Width           =   2595
   End
   Begin VB.TextBox TxtCNIC 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9863
      MaxLength       =   15
      TabIndex        =   11
      Top             =   6885
      Width           =   2070
   End
   Begin VB.TextBox TxtSectorName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   10568
      TabIndex        =   35
      Tag             =   "NC"
      Top             =   1905
      Width           =   1440
   End
   Begin VB.TextBox TxtSectorID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9503
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "C"
      Top             =   1905
      Width           =   705
   End
   Begin VB.TextBox TxtZoneName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   12008
      TabIndex        =   34
      Tag             =   "NC"
      Top             =   1905
      Width           =   1440
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
      TabIndex        =   30
      Top             =   960
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
         TabIndex        =   31
         Tag             =   "NC"
         Text            =   "DefInstallmentCustomers.frx":0ECA
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
         TabIndex        =   32
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.TextBox TxtPrefix 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   7883
      MaxLength       =   2
      TabIndex        =   26
      Tag             =   "NC"
      Top             =   1905
      Width           =   525
   End
   Begin VB.TextBox TxtID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8408
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1905
      Width           =   825
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7883
      MaxLength       =   30
      TabIndex        =   2
      Top             =   2550
      Width           =   5265
   End
   Begin VB.TextBox TxtFilter 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2798
      MaxLength       =   30
      TabIndex        =   20
      Tag             =   "NC"
      Top             =   2355
      Width           =   4395
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5700
      Left            =   1913
      TabIndex        =   21
      Top             =   2685
      Width           =   5295
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
      stylesets(0).Picture=   "DefInstallmentCustomers.frx":0F55
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
      Columns(0).Caption=   "Customer ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   6456
      Columns(1).Caption=   "Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   9340
      _ExtentY        =   10054
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
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   2865
      TabIndex        =   22
      Top             =   9825
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
      MICON           =   "DefInstallmentCustomers.frx":0F71
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   4185
      TabIndex        =   23
      Top             =   9825
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
      MICON           =   "DefInstallmentCustomers.frx":0F8D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   5505
      TabIndex        =   24
      Top             =   9825
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
      MICON           =   "DefInstallmentCustomers.frx":0FA9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   8310
      TabIndex        =   17
      Top             =   9825
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
      MICON           =   "DefInstallmentCustomers.frx":0FC5
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   9630
      TabIndex        =   18
      Top             =   9825
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
      MICON           =   "DefInstallmentCustomers.frx":0FE1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   10950
      TabIndex        =   19
      Top             =   9825
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
      MICON           =   "DefInstallmentCustomers.frx":0FFD
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSector 
      Height          =   330
      Left            =   10208
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1905
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
      MICON           =   "DefInstallmentCustomers.frx":1019
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpDateofBirth 
      Height          =   315
      Left            =   11933
      TabIndex        =   12
      Tag             =   "NC"
      Top             =   6885
      Width           =   1395
      _Version        =   65543
      _ExtentX        =   2461
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   16777215
      Enabled         =   0   'False
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
   Begin JeweledBut.JeweledButton BtnRefer1 
      Height          =   330
      Left            =   8595
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   8340
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
      MICON           =   "DefInstallmentCustomers.frx":1035
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnRefer2 
      Height          =   330
      Left            =   8595
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   8970
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
      MICON           =   "DefInstallmentCustomers.frx":1051
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnAddRefer1 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   13005
      TabIndex        =   62
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   8340
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      TX              =   "+"
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
      MICON           =   "DefInstallmentCustomers.frx":106D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnAddRefer2 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   13005
      TabIndex        =   63
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   8970
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      TX              =   "+"
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
      MICON           =   "DefInstallmentCustomers.frx":1089
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Installment"
      Height          =   195
      Left            =   9855
      TabIndex        =   65
      Top             =   7380
      Width           =   1230
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Promise Date"
      Height          =   195
      Left            =   7875
      TabIndex        =   64
      Top             =   7380
      Width           =   945
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      Height          =   195
      Left            =   10980
      TabIndex        =   61
      Top             =   8760
      Width           =   915
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refer ID 2"
      Height          =   195
      Left            =   7890
      TabIndex        =   60
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   8955
      TabIndex        =   59
      Top             =   8760
      Width           =   420
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      Height          =   195
      Left            =   10980
      TabIndex        =   55
      Top             =   8130
      Width           =   915
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refer ID 1"
      Height          =   195
      Left            =   7890
      TabIndex        =   54
      Top             =   8130
      Width           =   735
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   8955
      TabIndex        =   53
      Top             =   8130
      Width           =   420
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Nos"
      Height          =   195
      Left            =   7883
      TabIndex        =   49
      Top             =   6015
      Width           =   795
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      Height          =   195
      Left            =   12158
      TabIndex        =   48
      Top             =   6675
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address 1"
      Height          =   195
      Left            =   7883
      TabIndex        =   46
      Top             =   3555
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Left            =   7883
      TabIndex        =   45
      Top             =   5430
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No"
      Height          =   195
      Left            =   7883
      TabIndex        =   44
      Top             =   6675
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      Height          =   195
      Left            =   7883
      TabIndex        =   43
      Top             =   2955
      Width           =   915
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address 2"
      Height          =   195
      Left            =   7883
      TabIndex        =   42
      Top             =   4485
      Width           =   705
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cast"
      Height          =   195
      Left            =   10583
      TabIndex        =   41
      Top             =   5430
      Width           =   315
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNIC No"
      Height          =   195
      Left            =   9863
      TabIndex        =   40
      Top             =   6675
      Width           =   630
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector Name"
      Height          =   195
      Left            =   10568
      TabIndex        =   39
      Top             =   1695
      Width           =   930
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector ID"
      Height          =   195
      Left            =   9503
      TabIndex        =   38
      Top             =   1695
      Width           =   675
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zone Name"
      Height          =   195
      Left            =   12008
      TabIndex        =   37
      Top             =   1695
      Width           =   840
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
      Left            =   12698
      TabIndex        =   33
      Top             =   1515
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Installment Customers"
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
      TabIndex        =   29
      Top             =   270
      Width           =   2970
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      Height          =   195
      Left            =   7883
      TabIndex        =   28
      Top             =   1695
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   7883
      TabIndex        =   27
      Top             =   2340
      Width           =   420
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   500.533
      X2              =   499.533
      Y1              =   156
      Y2              =   562
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   2183
      TabIndex        =   25
      Top             =   2445
      Width           =   510
   End
End
Attribute VB_Name = "DefInstallmentCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vMode As FormMode
Dim vMaxBinID As Integer
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.

Private Sub BtnAddRefer1_Click()
   DefReferences.Show
End Sub

Private Sub BtnAddRefer2_Click()
   DefReferences.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
     If ActiveControl.Name = Grid.Name Then Call Grid_DblClick: Exit Sub
     keybd_event 9, 1, 1, 1
     KeyCode = 0
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtSectorID.Name: If FunSelectSector(ssFunctionKey, False) = True Then TxtName.SetFocus Else TxtSectorID.SetFocus
         Case TxtReferID1.Name: If FunSelectRefer1(ssFunctionKey, False) = True Then TxtReferID2.SetFocus Else TxtReferID1.SetFocus
         Case TxtReferID2.Name: If FunSelectRefer2(ssFunctionKey, False) = True Then If BtnSave.Enabled Then BtnSave.SetFocus Else TxtReferID2.SetFocus
      End Select
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
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
         Case vbKeyH
               FraHelp.ZOrder 0
               FraHelp.Visible = True
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
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   End If
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
      TxtFilter.Text = Chr(KeyAscii): TxtFilter.SelStart = Len(TxtFilter.Text):  TxtFilter.SetFocus
   End Select
End Sub

Private Sub BtnClear_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ",Null,'Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  FormStatus = SelectionMode
End Sub

Private Sub BtnClose_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ",Null,'Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniInstallmentCustomers", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
  Dim vtbl As String
  If Rs.RecordCount > 0 Then
    If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
    Dim vid As String
    vid = Rs!PartyID
    vtbl = Common.ChildDataExists("Parties", "PartyId=" & vid, "") & Common.ChildDataExists("ChartoFAccounts", "AccountNo=" & vid, "Parties")
    If vtbl <> "" Then
      MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
      Exit Sub
    End If
    cn.BeginTrans
    
'    vMaxBinID = FunGetMaxBinID
    ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header----------------------------------------------
'    CN.Execute ("Insert Into Bin_Parties Select " & vMaxBinID & ",'" & Date & "',* from Parties Where PartyID = " & TxtPrefix.Text & TxtID.Text)
    
    Call ActivityLog("Customers", eDelete, , , TxtPrefix.Text & TxtID.Text)
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtPrefix.Text & TxtID.Text & ",Null,'Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Rs.Delete
    cn.Execute ("Delete From ChartOfAccounts Where AccountNo = " & vid)
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
   vUserAction = UserAuthentication("MniInstallmentCustomers", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
  If vIsNewRecord = False Then Call ActivityLog("Customers", eEdit, , , TxtPrefix.Text & TxtID.Text)
  Set Rs = New ADODB.Recordset
  Rs.Open " Select * FROM Parties where PartyID = " & Val(TxtPrefix.Text & TxtID.Text), cn, adOpenDynamic, adLockOptimistic
  
  Call UserActivities
  
  If vIsNewRecord Then
    cn.BeginTrans
    cn.Execute ("Insert into chartofaccounts(AccountNo,UserNo,AccountName,AccountType,AccountDepth,Narration,ParentAccountNo,OpeningDebit,OpeningCredit,IsDetailed,IsLocked,IsEditable,BalFlag,PLFlag,ExpFlag) values (" & _
    Val(TxtPrefix.Text & TxtID.Text) & ",1,'" & Replace(TxtName.Text, "'", "''") & "','Customers',2,'" & Replace(TxtAddress.Text, "'", "''") & "','62',0,0,1,0,1,0,' ',0)")
    Rs.AddNew
    Rs!PartyID = Val(TxtPrefix.Text & TxtID.Text)
    Rs!IsLockParty = 0
  Else
    cn.BeginTrans
    cn.Execute ("Update Chartofaccounts set Accountname = '" & Replace(TxtName.Text, "'", "''") & "',Narration = '" & Replace(TxtAddress.Text, "'", "''") & "' Where AccountNo = " & Rs!PartyID)
  End If
  Rs!SectorID = IIf(Trim(TxtSectorID.Text) = "", Null, TxtSectorID.Text)
  Rs!partyname = TxtName.Text
  Rs!FName = TxtFName.Text
  Rs!Address = TxtAddress.Text
  Rs!Address2 = TxtAddress2.Text
  Rs!City = TxtCity.Text
  Rs!Cast = TxtCast.Text
  Rs!Phone1 = TxtPhone1.Text
  Rs!Phone2 = TxtPhone2.Text
  Rs!Mobile = TxtMobileNo.Text
  Rs!DateofBirth = IIf(ChkDateofBirth.Value = 1, DtpDateofBirth.DateValue, Null)
  Rs!CNIC = TxtCNIC.Text
  Rs!PromiseDate = IIf(Trim(TxtPromiseDate.Text) = "", Null, TxtPromiseDate.Text)
  Rs!NoOfInstallments = IIf(Val(TxtNoOfInstallments.Text) = 0, Null, Val(TxtNoOfInstallments.Text))
  Rs!PartyType = "C"
  Rs!ReferID1 = IIf(Trim(TxtReferID1.Text) = "", Null, TxtReferID1.Text)
  Rs!ReferID2 = IIf(Trim(TxtReferID2.Text) = "", Null, TxtReferID2.Text)
  Rs.Update
  cn.CommitTrans
  Set Rs = New ADODB.Recordset
  Rs.Open "Select * FROM Parties Where PartyType = 'C'", cn, adOpenDynamic, adLockOptimistic
  If vIsNewRecord = True Then Call ActivityLog("Customers", eAdd, , , TxtPrefix.Text & TxtID.Text)
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
      MsgBox "Please specify a Customer ID", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
    If Not IsNumeric(TxtID.Text) Then
      MsgBox "The Customer ID must be numeric", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
    If Val(TxtNoOfInstallments.Text) <= 0 Or Val(TxtNoOfInstallments.Text) > 50 Then
      MsgBox "Please correct No. of Installmnent", vbExclamation, "Alert"
      If TxtNoOfInstallments.Enabled And TxtNoOfInstallments.Visible Then TxtNoOfInstallments.SetFocus
      Exit Function
    End If
    
    If Val(TxtPromiseDate.Text) <= 0 Or Val(TxtPromiseDate.Text) > 30 Then
      MsgBox "Please correct Promise Date", vbExclamation, "Alert"
      If TxtPromiseDate.Enabled And TxtPromiseDate.Visible Then TxtPromiseDate.SetFocus
      Exit Function
    End If
    
  End If
  If Trim(TxtName.Text) = "" Then
    MsgBox "Please specify a Customer Name", vbExclamation, "Alert"
    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
    Exit Function
  End If
  If TxtID.Enabled = True And cn.Execute("select count(*) from chartofaccounts where accountno = " & Val(TxtPrefix.Text & TxtID.Text)).Fields(0) > 0 Then
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

Private Sub ChkDateofBirth_Click()
   DtpDateofBirth.Enabled = IIf(ChkDateofBirth.Value = 1, True, False)
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
   SetWindowText Me.hWnd, "Customers"
   HelpLocation Me
   Set Rs = New ADODB.Recordset
   Rs.Open "Select * FROM Parties Where PartyType = 'C'", cn, adOpenDynamic, adLockOptimistic
   Grid.Columns("ID").DataField = "PartyId"
   Grid.Columns("Name").DataField = "PartyName"
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
      TxtPrefix.Text = "62"
      TxtID.Text = FunGetMaxID
      TxtFilter.Text = ""
      Grid.Enabled = False
      Set Grid.DataSource = Rs
      TxtFilter.Enabled = False
      If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
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
      TxtName.SetFocus
      TxtFilter.Text = ""
      vIsNewRecord = False
    Case Is = ChangeMode
      BtnSave.Enabled = True
    Case Is = SelectionMode
      Grid.Enabled = True
      Call SubClearFields(False)
      Call Grid_RowColChange(0, 0)
      TxtFilter.Enabled = True
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
    TxtID.Text = Mid(Grid.Columns("ID").Text, 3)
    TxtName.Text = Grid.Columns("Name").Text
    TxtFName.Text = IIf(IsNull(Rs!FName), "", Rs!FName)
    TxtAddress.Text = IIf(IsNull(Rs!Address), "", Rs!Address)
    TxtAddress2.Text = IIf(IsNull(Rs!Address2), "", Rs!Address2)
    TxtCity.Text = IIf(IsNull(Rs!City), "", Rs!City)
    TxtCast.Text = IIf(IsNull(Rs!Cast), "", Rs!Cast)
    TxtPhone1.Text = IIf(IsNull(Rs!Phone1), "", Rs!Phone1)
    TxtPhone2.Text = IIf(IsNull(Rs!Phone2), "", Rs!Phone2)
    TxtMobileNo.Text = IIf(IsNull(Rs!Mobile), "", Rs!Mobile)
    TxtCNIC.Text = IIf(IsNull(Rs!CNIC), "", Rs!CNIC)
    TxtPromiseDate.Text = IIf(IsNull(Rs!PromiseDate), "", Rs!PromiseDate)
    TxtNoOfInstallments.Text = IIf(IsNull(Rs!NoOfInstallments), "", Rs!NoOfInstallments)
    
    If IsNull(Rs!DateofBirth) Then
      DtpDateofBirth.DateValue = ""
      ChkDateofBirth.Value = 0
    Else
      DtpDateofBirth.DateValue = Rs!DateofBirth
      ChkDateofBirth.Value = 1
    End If
    TxtReferID1.Text = IIf(IsNull(Rs!ReferID1), "", Rs!ReferID1)
    If TxtReferID1.Text = "" Then
      TxtReferName1.Text = ""
      TxtReferFName1.Text = ""
    Else
      With cn.Execute("select * from Refers where ReferID =" & Val(TxtReferID1.Text))
        If .RecordCount > 0 Then
           TxtReferName1.Text = !Name
           TxtReferFName1.Text = !FName
        End If
      End With
    End If
    TxtReferID2.Text = IIf(IsNull(Rs!ReferID2), "", Rs!ReferID2)
    If TxtReferID2.Text = "" Then
      TxtReferName2.Text = ""
      TxtReferFName2.Text = ""
    Else
      With cn.Execute("select * from Refers where ReferID =" & Val(TxtReferID2.Text))
        If .RecordCount > 0 Then
           TxtReferName2.Text = !Name
           TxtReferFName2.Text = !FName
        End If
      End With
    End If
    TxtSectorID.Text = IIf(IsNull(Rs!SectorID), "", Rs!SectorID)
    If TxtSectorID.Text = "" Then
      TxtSectorName.Text = ""
      TxtZoneName.Text = ""
    Else
      With cn.Execute("select * from sectors s inner join Zones t on s.ZoneID = t.ZoneID where sectorid =" & Val(TxtSectorID.Text))
        If .RecordCount > 0 Then
           TxtSectorName.Text = !SectorName
           TxtZoneName.Text = !ZoneName
        End If
      End With
    End If
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
      ElseIf TypeOf ctl Is ComboBox Then
         ctl.Enabled = Enable
      End If
   Next
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As String
  FunGetMaxID = cn.Execute("Select isnull(max(cast(substring(cast(accountno as varchar(10)),3,10) as int)),0) + 1 from chartofaccounts Where AccountNo like '62%' and isdetailed=1").Fields(0)
End Function

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtAddress_LostFocus()
   TxtAddress.Text = StrConv(TxtAddress.Text, vbProperCase)
End Sub

Private Sub TxtCity_LostFocus()
   TxtCity.Text = StrConv(TxtCity.Text, vbProperCase)
End Sub

Private Sub TxtFilter_Change()
   On Error GoTo ErrorHandler
   'If Me.ActiveControl.Name <> TxtFilter.Name Then Exit Sub
   'If Trim(TxtFilter.Text) = "" Then Grid.MoveFirst: Exit Sub
   Set Rs = New ADODB.Recordset
   Rs.Open " Select * FROM Parties where PartyType = 'C' and PartyName like '%" & Replace(TxtFilter.Text, "'", "''") & "%' Order by PartyName", cn, adOpenDynamic, adLockOptimistic
   Set Grid.DataSource = Rs
   'Rs.Find "PartyName like '" & Replace(TxtFilter.Text, "'", "''") & "%'", , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtMobileNo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtCNIC_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtName_Change()
   If Me.ActiveControl.Name <> TxtName.Name Then Exit Sub
   TxtFilter.Text = TxtName.Text
End Sub

Private Sub TxtName_LostFocus()
   TxtName.Text = StrConv(TxtName.Text, vbProperCase)
End Sub

Private Sub TxtFName_LostFocus()
   TxtFName.Text = StrConv(TxtFName.Text, vbProperCase)
End Sub

Private Sub TxtNoOfInstallments_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtPhone1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtPhone2_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxBinID = cn.Execute("Select isnull(max(BinID),0)+1 from Bin_Parties ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
    If vIsNewRecord = False Then
        If TxtName.Text <> IIf(IsNull(Rs!partyname), "", Rs!partyname) Then
            cn.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated Customer-" & Rs!partyname & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtAddress.Text <> IIf(IsNull(Rs!Address), "", Rs!Address) Then
            cn.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated Address-" & Rs!Address & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtCity.Text <> IIf(IsNull(Rs!City), "", Rs!City) Then
            cn.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated City-" & Rs!City & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtPhone1.Text <> IIf(IsNull(Rs!Phone1), "", Rs!Phone1) Then
            cn.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated Phone1-" & Rs!Phone1 & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtPhone2.Text <> IIf(IsNull(Rs!Phone2), "", Rs!Phone2) Then
            cn.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated Phone2-" & Rs!Phone2 & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtMobileNo.Text <> IIf(IsNull(Rs!Mobile), "", Rs!Mobile) Then
            cn.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null , 'Updated Mobile-" & Rs!Mobile & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
   Else
        cn.Execute ("Insert Into UserActivities values ('Customers'" & "," & TxtID.Text & ", Null ,'Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

Private Sub BtnSector_Click()
   If FunSelectSector(ssButton, False) = True Then
     If TxtName.Enabled Then TxtName.SetFocus
   Else
      TxtSectorID.SetFocus
   End If
End Sub

Private Sub TxtPromiseDate_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
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
   If TxtSectorName.Text <> "" Then Exit Sub
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

Private Function FunSelectSector(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSector.Show vbModal, Me
        If SchSector.ParaOutSectorID = "" Then FunSelectSector = False: Exit Function
        TxtSectorID.Text = SchSector.ParaOutSectorID
    End If
    '---------------------------
    vStrSQL = "Select * FROM Sectors s inner join Zones t on t.ZoneID = s.ZoneID where SectorID=" & Val(TxtSectorID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSectorName.Text = !SectorName
          TxtZoneName.Text = !ZoneName
          FunSelectSector = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectSector = False
          .Close
          TxtSectorID.Text = ""
          TxtSectorName.Text = ""
          TxtZoneName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectRefer1(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchRefer.Show vbModal, Me
        If SchRefer.ParaOutReferID = "" Then FunSelectRefer1 = False: Exit Function
        TxtReferID1.Text = SchRefer.ParaOutReferID
    End If
    '---------------------------
    vStrSQL = "Select * FROM Refers where ReferID=" & Val(TxtReferID1.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtReferName1.Text = !Name
          TxtReferFName1.Text = !FName
          FunSelectRefer1 = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectRefer1 = False
          .Close
          TxtReferID1.Text = ""
          TxtReferName1.Text = ""
          TxtReferFName1.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnRefer1_Click()
   If FunSelectRefer1(ssButton, False) = True Then
      If TxtReferID2.Enabled Then TxtReferID2.SetFocus
   Else
      If TxtReferID1.Enabled Then TxtReferID1.SetFocus
   End If
End Sub

Private Sub TxtReferID1_Change()
   If TxtReferID1.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtReferID1.Name Then Exit Sub
   If TxtReferName1.Text <> "" Then TxtReferName1.Text = ""
   If TxtReferFName1.Text <> "" Then TxtReferFName1.Text = ""
End Sub

Private Sub TxtReferID1_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtReferID1.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtReferID1.Text = "" Then Exit Sub
   If TxtReferName1.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectRefer1(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectRefer1(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectRefer2(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchRefer.Show vbModal, Me
        If SchRefer.ParaOutReferID = "" Then FunSelectRefer2 = False: Exit Function
        TxtReferID2.Text = SchRefer.ParaOutReferID
    End If
    '---------------------------
    vStrSQL = "Select * FROM Refers where ReferID=" & Val(TxtReferID2.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtReferName2.Text = !Name
          TxtReferFName2.Text = !FName
          FunSelectRefer2 = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectRefer2 = False
          .Close
          TxtReferID2.Text = ""
          TxtReferName2.Text = ""
          TxtReferFName2.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnRefer2_Click()
   If FunSelectRefer2(ssButton, False) = True Then
      If BtnSave.Enabled Then BtnSave.SetFocus
   Else
      If TxtReferID2.Enabled Then TxtReferID2.SetFocus
   End If
End Sub

Private Sub TxtReferID2_Change()
   If TxtReferID2.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtReferID2.Name Then Exit Sub
   If TxtReferName2.Text <> "" Then TxtReferName2.Text = ""
   If TxtReferFName2.Text <> "" Then TxtReferFName2.Text = ""
End Sub

Private Sub TxtReferID2_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtReferID2.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtReferID2.Text = "" Then Exit Sub
   If TxtReferName2.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectRefer2(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectRefer2(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

