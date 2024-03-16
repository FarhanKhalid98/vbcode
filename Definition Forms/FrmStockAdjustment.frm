VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{CA42C609-3527-11D8-8B1B-004095005536}#1.0#0"; "CipherAGB.ocx"
Begin VB.Form FrmStockAdjustment 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmStockAdjustment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   765
      TabIndex        =   105
      Top             =   5805
      Visible         =   0   'False
      Width           =   2295
      Begin SITextBox.Txt TxtSerial 
         Height          =   315
         Left            =   120
         TabIndex        =   106
         Top             =   240
         Width           =   2040
         _ExtentX        =   3598
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
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridSerial 
         Height          =   1500
         Left            =   120
         TabIndex        =   107
         Top             =   555
         Width           =   2040
         ScrollBars      =   2
         _Version        =   196616
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   2
         stylesets.count =   1
         stylesets(0).Name=   "SelectedRow"
         stylesets(0).ForeColor=   -2147483634
         stylesets(0).BackColor=   -2147483635
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
         stylesets(0).Picture=   "FrmStockAdjustment.frx":0ECA
         AllowDelete     =   -1  'True
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
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "ProductID"
         Columns(0).Name =   "ProductID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3096
         Columns(1).Caption=   "Serial No"
         Columns(1).Name =   "Serial"
         Columns(1).CaptionAlignment=   2
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   3598
         _ExtentY        =   2646
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
   End
   Begin CipherAGBLib.CipherAGB CipherAGB1 
      Height          =   240
      Left            =   12308
      TabIndex        =   100
      Top             =   1545
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
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
      Height          =   3840
      Left            =   13800
      TabIndex        =   51
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
         Height          =   3315
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   52
         Tag             =   "NC"
         Text            =   "FrmStockAdjustment.frx":0EE6
         Top             =   390
         Width           =   3885
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
         TabIndex        =   53
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.ComboBox CmbPackName 
      Height          =   315
      Left            =   5723
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3900
      Width           =   1665
   End
   Begin SITextBox.Txt TxtAdjID 
      Height          =   315
      Left            =   1650
      TabIndex        =   0
      Top             =   1245
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   9
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
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9031
      TabIndex        =   17
      Top             =   9390
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
      MICON           =   "FrmStockAdjustment.frx":0FC6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7711
      TabIndex        =   13
      Top             =   9390
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
      MICON           =   "FrmStockAdjustment.frx":0FE2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5071
      TabIndex        =   15
      Top             =   9390
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Open"
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
      MICON           =   "FrmStockAdjustment.frx":0FFE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10351
      TabIndex        =   18
      Top             =   9390
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
      MICON           =   "FrmStockAdjustment.frx":101A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6391
      TabIndex        =   14
      Top             =   9390
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
      MICON           =   "FrmStockAdjustment.frx":1036
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtDifferenceAmount 
      Height          =   315
      Left            =   12308
      TabIndex        =   21
      Top             =   8550
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
   End
   Begin SITextBox.Txt TxtCost 
      Height          =   315
      Left            =   11348
      TabIndex        =   23
      Top             =   3900
      Width           =   660
      _ExtentX        =   1164
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
      Height          =   3840
      Left            =   1733
      TabIndex        =   24
      Top             =   4215
      Width           =   11895
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   15
      stylesets.count =   1
      stylesets(0).Name=   "Select"
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
      stylesets(0).Picture=   "FrmStockAdjustment.frx":1052
      AllowUpdate     =   0   'False
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
      SelectTypeRow   =   1
      RowNavigation   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   15
      Columns(0).Width=   3122
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3889
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "ProductName"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2963
      Columns(2).Caption=   "Pack Name"
      Columns(2).Name =   "PackName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   953
      Columns(3).Caption=   "Pack"
      Columns(3).Name =   "Pack"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1005
      Columns(4).Caption=   "Qty(P)"
      Columns(4).Name =   "SQtyPack"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   4
      Columns(4).FieldLen=   256
      Columns(5).Width=   1005
      Columns(5).Caption=   "Qty(L)"
      Columns(5).Name =   "SQtyLoose"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   4
      Columns(5).FieldLen=   256
      Columns(6).Width=   1005
      Columns(6).Caption=   "Qty(P)"
      Columns(6).Name =   "OQtyPack"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   4
      Columns(6).FieldLen=   256
      Columns(7).Width=   1005
      Columns(7).Caption=   "Qty(L)"
      Columns(7).Name =   "OQtyLoose"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   4
      Columns(7).FieldLen=   256
      Columns(8).Width=   1005
      Columns(8).Caption=   "Qty(P)"
      Columns(8).Name =   "UQtyPack"
      Columns(8).Alignment=   1
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1005
      Columns(9).Caption=   "Qty(L)"
      Columns(9).Name =   "UQtyLoose"
      Columns(9).Alignment=   1
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1164
      Columns(10).Caption=   "Cost"
      Columns(10).Name=   "Cost"
      Columns(10).Alignment=   1
      Columns(10).CaptionAlignment=   2
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   4
      Columns(10).FieldLen=   256
      Columns(11).Width=   2355
      Columns(11).Caption=   "Amount"
      Columns(11).Name=   "Amount"
      Columns(11).Alignment=   1
      Columns(11).CaptionAlignment=   2
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   5
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "PackingID"
      Columns(12).Name=   "PackingID"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "ProductID"
      Columns(13).Name=   "ProductID"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   3200
      Columns(14).Caption=   "IsSerial"
      Columns(14).Name=   "IsSerial"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20981
      _ExtentY        =   6773
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpAdjustmentDate 
      Height          =   315
      Left            =   2940
      TabIndex        =   1
      Top             =   1245
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   4320
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   1245
      Width           =   675
      _ExtentX        =   1191
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
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   5355
      TabIndex        =   28
      Tag             =   "NC"
      Top             =   1245
      Width           =   2340
      _ExtentX        =   4128
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
      Left            =   4995
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1245
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
      MICON           =   "FrmStockAdjustment.frx":106E
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtMultiplier 
      Height          =   315
      Left            =   7388
      TabIndex        =   6
      Top             =   3900
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   556
      Alignment       =   1
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
   Begin SITextBox.Txt TxtSQtyLoose 
      Height          =   315
      Left            =   8498
      TabIndex        =   8
      Top             =   3900
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   9
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
      DecimalPoint    =   3
      IntegralPoint   =   5
      MandatoryBackColor=   14473725
   End
   Begin SITextBox.Txt TxtSQtyPack 
      Height          =   315
      Left            =   7928
      TabIndex        =   7
      Top             =   3900
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Alignment       =   1
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
      MandatoryBackColor=   14473725
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   12008
      TabIndex        =   36
      Top             =   3900
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Alignment       =   1
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   1733
      TabIndex        =   4
      Top             =   3900
      Width           =   1410
      _ExtentX        =   2487
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
      IntegralPoint   =   15
   End
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   3143
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3900
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
      MICON           =   "FrmStockAdjustment.frx":108A
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   3503
      TabIndex        =   42
      Top             =   3900
      Width           =   2220
      _ExtentX        =   3916
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
   Begin SITextBox.Txt TxtOQtyLoose 
      Height          =   315
      Left            =   9638
      TabIndex        =   10
      Top             =   3900
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   9
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
      DecimalPoint    =   3
      IntegralPoint   =   5
      MandatoryBackColor=   14679293
   End
   Begin SITextBox.Txt TxtOQtyPack 
      Height          =   315
      Left            =   9068
      TabIndex        =   9
      Top             =   3900
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
      MandatoryBackColor=   14679293
   End
   Begin SITextBox.Txt TxtUQtyLoose 
      Height          =   315
      Left            =   10778
      TabIndex        =   12
      Top             =   3900
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   9
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
      DecimalPoint    =   3
      IntegralPoint   =   5
      MandatoryBackColor=   14679293
   End
   Begin SITextBox.Txt TxtUQtyPack 
      Height          =   315
      Left            =   10208
      TabIndex        =   11
      Top             =   3900
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
      MandatoryBackColor=   14679293
   End
   Begin SITextBox.Txt TxtPID 
      Height          =   315
      Left            =   11228
      TabIndex        =   55
      Top             =   1515
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
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
   Begin SITextBox.Txt TxtTotUQtyLoose 
      Height          =   315
      Left            =   9878
      TabIndex        =   57
      Top             =   8550
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
   End
   Begin SITextBox.Txt TxtTotUQtyPack 
      Height          =   315
      Left            =   8858
      TabIndex        =   59
      Top             =   8550
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
   End
   Begin SITextBox.Txt TxtTotOQtyLoose 
      Height          =   315
      Left            =   6413
      TabIndex        =   62
      Top             =   8550
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
   End
   Begin SITextBox.Txt TxtTotOQtyPack 
      Height          =   315
      Left            =   5303
      TabIndex        =   64
      Top             =   8550
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
   End
   Begin SITextBox.Txt TxtTotSQtyLoose 
      Height          =   315
      Left            =   2813
      TabIndex        =   67
      Top             =   8550
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
   End
   Begin SITextBox.Txt TxtTotSQtyPack 
      Height          =   315
      Left            =   1703
      TabIndex        =   69
      Top             =   8550
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpTransectionDate 
      Height          =   315
      Left            =   7545
      TabIndex        =   3
      Top             =   2025
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
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin SITextBox.Txt TxtTotSAmount 
      Height          =   315
      Left            =   3983
      TabIndex        =   77
      Top             =   8550
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
   End
   Begin SITextBox.Txt TxtTotOAmount 
      Height          =   315
      Left            =   7538
      TabIndex        =   79
      Top             =   8550
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
   End
   Begin SITextBox.Txt TxtTotUAmount 
      Height          =   315
      Left            =   11003
      TabIndex        =   81
      Top             =   8550
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
   End
   Begin SITextBox.Txt TxtBillID 
      Height          =   315
      Left            =   3330
      TabIndex        =   85
      Tag             =   "NC"
      Top             =   2025
      Width           =   720
      _ExtentX        =   1270
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
   Begin SITextBox.Txt TxtSaleReturnID 
      Height          =   315
      Left            =   4050
      TabIndex        =   86
      Tag             =   "NC"
      Top             =   2025
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
      Masked          =   1
      IntegralPoint   =   10
   End
   Begin SITextBox.Txt TxtPurID 
      Height          =   315
      Left            =   1665
      TabIndex        =   87
      Tag             =   "NC"
      Top             =   2025
      Width           =   675
      _ExtentX        =   1191
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
   Begin SITextBox.Txt TxtPurReturnID 
      Height          =   315
      Left            =   2340
      TabIndex        =   88
      Tag             =   "NC"
      Top             =   2025
      Width           =   990
      _ExtentX        =   1746
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
   Begin SITextBox.Txt TxtTransferID 
      Height          =   315
      Left            =   5070
      TabIndex        =   89
      Tag             =   "NC"
      Top             =   2025
      Width           =   990
      _ExtentX        =   1746
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
   Begin SITextBox.Txt TxtManufacturedID 
      Height          =   315
      Left            =   6060
      TabIndex        =   90
      Tag             =   "NC"
      Top             =   2025
      Width           =   1485
      _ExtentX        =   2619
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
   Begin JeweledBut.JeweledButton BtnPrint 
      Height          =   420
      Left            =   3751
      TabIndex        =   16
      Top             =   9390
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Print"
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
      MICON           =   "FrmStockAdjustment.frx":10A6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnLoadData 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   12263
      TabIndex        =   91
      Top             =   3165
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Load Data"
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
      MICON           =   "FrmStockAdjustment.frx":10C2
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPortNumber 
      Height          =   315
      Left            =   11543
      TabIndex        =   92
      Tag             =   "NC"
      Top             =   3210
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   2
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
   Begin JeweledBut.JeweledButton BtnRefresh 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   12263
      TabIndex        =   94
      Top             =   2760
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Refresh"
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
      MICON           =   "FrmStockAdjustment.frx":10DE
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   7740
      TabIndex        =   95
      Tag             =   "NC"
      Top             =   1245
      Width           =   945
      _ExtentX        =   1667
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
   Begin SITextBox.Txt TxtOrganizationName 
      Height          =   315
      Left            =   9045
      TabIndex        =   96
      Tag             =   "NC"
      Top             =   1245
      Width           =   1980
      _ExtentX        =   3493
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
      Left            =   8685
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   1245
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
      MICON           =   "FrmStockAdjustment.frx":10FA
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnProductRange 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2790
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   3555
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
      MICON           =   "FrmStockAdjustment.frx":1116
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtRemarks 
      Height          =   315
      Left            =   1665
      TabIndex        =   103
      Top             =   2820
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   200
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
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Left            =   1665
      TabIndex        =   104
      Top             =   2610
      Width           =   630
   End
   Begin VB.Label LblPre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   1823
      TabIndex        =   101
      Top             =   3345
      Visible         =   0   'False
      Width           =   6120
   End
   Begin VB.Label LblOrganizationID 
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
      Left            =   7740
      TabIndex        =   99
      Top             =   1050
      Width           =   1335
   End
   Begin VB.Label LblOrganizationName 
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
      Left            =   9135
      TabIndex        =   98
      Top             =   1050
      Width           =   1620
   End
   Begin VB.Label LblPort 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Port Number"
      Height          =   195
      Left            =   11318
      TabIndex        =   93
      Top             =   2985
      Width           =   885
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manufactured ID"
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
      Left            =   6075
      TabIndex        =   84
      Top             =   1815
      Width           =   1440
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer ID"
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
      Left            =   5085
      TabIndex        =   83
      Top             =   1815
      Width           =   975
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   11003
      TabIndex        =   82
      Top             =   8340
      Width           =   645
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   7538
      TabIndex        =   80
      Top             =   8340
      Width           =   645
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   3983
      TabIndex        =   78
      Top             =   8340
      Width           =   645
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "P Return ID"
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
      Left            =   2295
      TabIndex        =   76
      Top             =   1815
      Width           =   1020
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transection Date"
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
      Left            =   7515
      TabIndex        =   75
      Top             =   1815
      Width           =   1485
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "S Return ID"
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
      Left            =   4050
      TabIndex        =   74
      Top             =   1815
      Width           =   1020
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale ID"
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
      Left            =   3375
      TabIndex        =   73
      Top             =   1815
      Width           =   645
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur ID"
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
      Left            =   1665
      TabIndex        =   72
      Top             =   1815
      Width           =   555
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "---- Set Stock ----"
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
      Left            =   2633
      TabIndex        =   71
      Top             =   8115
      Width           =   1455
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty (L)"
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
      Left            =   1688
      TabIndex        =   70
      Top             =   8310
      Width           =   1080
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty (L)"
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
      Left            =   2858
      TabIndex        =   68
      Top             =   8340
      Width           =   1080
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "---- Over Stock ----"
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
      Left            =   5918
      TabIndex        =   66
      Top             =   8115
      Width           =   1575
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty (L)"
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
      Left            =   5303
      TabIndex        =   65
      Top             =   8310
      Width           =   1080
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty (L)"
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
      Left            =   6413
      TabIndex        =   63
      Top             =   8340
      Width           =   1080
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "---- Under Stock ----"
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
      Left            =   9458
      TabIndex        =   61
      Top             =   8100
      Width           =   1680
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty (L)"
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
      Left            =   8753
      TabIndex        =   60
      Top             =   8340
      Width           =   1080
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty (L)"
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
      Left            =   9878
      TabIndex        =   58
      Top             =   8340
      Width           =   1080
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   11273
      TabIndex        =   56
      Top             =   1320
      Visible         =   0   'False
      Width           =   720
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
      Left            =   12878
      TabIndex        =   54
      Top             =   1650
      Width           =   435
   End
   Begin VB.Label LblQtyPack 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   12308
      TabIndex        =   50
      Top             =   2100
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "- Set Stock -"
      Height          =   195
      Left            =   7988
      TabIndex        =   49
      Top             =   3510
      Width           =   885
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
      Height          =   195
      Left            =   10208
      TabIndex        =   48
      Top             =   3705
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
      Height          =   195
      Left            =   10778
      TabIndex        =   47
      Top             =   3705
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "- Under Stock -"
      Height          =   195
      Left            =   10178
      TabIndex        =   46
      Top             =   3510
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
      Height          =   195
      Left            =   9638
      TabIndex        =   45
      Top             =   3705
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
      Height          =   195
      Left            =   9068
      TabIndex        =   44
      Top             =   3705
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "- Over Stock -"
      Height          =   195
      Left            =   9038
      TabIndex        =   43
      Top             =   3510
      Width           =   990
   End
   Begin VB.Label LblStock 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   9360
      TabIndex        =   40
      Top             =   2040
      Width           =   1035
   End
   Begin VB.Label LblStockCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   9360
      TabIndex        =   39
      Top             =   1725
      Width           =   720
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Adjustment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   38
      Top             =   270
      Width           =   3060
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   11993
      TabIndex        =   37
      Top             =   3705
      Width           =   540
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (P)"
      Height          =   195
      Left            =   7928
      TabIndex        =   35
      Top             =   3705
      Width           =   480
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty (L)"
      Height          =   195
      Left            =   8498
      TabIndex        =   34
      Top             =   3705
      Width           =   465
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   5723
      TabIndex        =   33
      Top             =   3705
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   195
      Left            =   7388
      TabIndex        =   32
      Top             =   3705
      Width           =   375
   End
   Begin VB.Label LblStoreName 
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
      Left            =   5355
      TabIndex        =   31
      Top             =   1050
      Width           =   1005
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
      Left            =   4320
      TabIndex        =   30
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   195
      Left            =   11348
      TabIndex        =   27
      Top             =   3705
      Width           =   315
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   1733
      TabIndex        =   26
      Top             =   3705
      Width           =   375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3503
      TabIndex        =   25
      Top             =   3705
      Width           =   1020
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Difference Amount"
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
      Left            =   12083
      TabIndex        =   22
      Top             =   8355
      Width           =   1590
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adj Date"
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
      Left            =   2955
      TabIndex        =   20
      Top             =   1005
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Adj ID"
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
      Left            =   1665
      TabIndex        =   19
      Top             =   1005
      Width           =   540
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmStockAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vCounter, vGridRows As Integer
Dim vUnitPrice As Double
Dim RsReport As New ADODB.Recordset
Dim RsBody As New ADODB.Recordset
Dim RsBodySerial As New ADODB.Recordset
Dim Flag As Boolean
Dim ssql As String, vAutoEnterBeforeQty, vIsSerial  As Boolean
Dim vQtyLoose As Double
Dim vStrComp, vRandomID As String, vCompanyName As String, vAddress As String, vEmail As String, vStrSQL
Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long

'----------------------------------

Private Sub SubCalculateBody()
   On Error GoTo ErrorHandler
   'TxtAmount.Text = SelfRound(Val(vUnitPrice) * (Val(TxtSQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtSQtyLoose.Text)))
   TxtAmount.Text = Round(Val(vUnitPrice) * (Val(TxtUQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtUQtyLoose.Text)), 2) + Round(Val(vUnitPrice) * (Val(TxtOQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtOQtyLoose.Text)), 2)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCalculateHeader()
   On Error GoTo ErrorHandler
   TxtDifferenceAmount.Text = Val(TxtTotOAmount.Text) - Val(TxtTotUAmount.Text)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

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
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectStore = False
          .Close
          TxtStoreID.Text = ""
          TxtStoreName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
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
      SchProduct.ParaInWhere = " and isLocked = 0 and isRawProduct = 0 and (StoreID is Null or StoreID = " & TxtStoreID.Text & ")"
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
   If TxtCode.Enabled = False Then FunSelectProduct = False: Exit Function
   If Trim(TxtCode.Text) = "" Then FunSelectProduct = False: Exit Function
   If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
   vStrSQL = " SELECT p.productid, Code, ProductName, PackingName, IsSerial, isnull(Multiplier,0) as Multiplier, round(PurPrice/isnull(Multiplier,1),4) as Cost" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "p.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')"
           '+ " where p.productid = " & Val(TxtCode.Text) & " or Code='" & TxtCode.Text & "'"
  
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtPID.Text = !Productid
         TxtProductName.Text = !ProductName
         vIsSerial = !IsSerial
         TxtMultiplier.Text = !Multiplier
         TxtSQtyLoose.Text = IIf(Val(TxtMultiplier.Text) > 0, "", IIf(Val(TxtSQtyLoose.Text) = 0, 1, TxtSQtyLoose.Text))
         vUnitPrice = !Cost
         TxtCost.Text = !Cost
         If Not IsNull(!PackingName) Then CmbPackName.Text = !PackingName Else CmbPackName.ListIndex = 0
'         With CN.Execute("select Cost from CurrentStock where productid ='" & TxtPID.Text & "'")
'            If .RecordCount > 0 Then
'               vUnitPrice = !Cost
'               TxtCost.Text = !Cost
'            Else
'               TxtCost.Text = ""
'            End If
'         End With
'         VStrSQL = "select isnull(dbo.FunStockOrg(" & IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text) & ",'" & TxtPID.Text & "'," & TxtStoreID.Text & "," & Val(TxtPurID.Text) & "," & Val(TxtPurReturnID.Text) & "," & Val(TxtBillID.Text) & "," & Val(TxtSaleReturnID.Text) & "," & Val(TxtTransferID.Text) & "," & Val(TxtManufacturedID.Text) & ",'" & DtpTransectionDate.DateValue & "'," & Val(TxtAdjID.Text) & "),0)"
'         vQtyLoose = CN.Execute(VStrSQL).Fields(0).Value
'         LblStock.Caption = vQtyLoose

'         vStrSQL = "select isnull(dbo.FunStockOrg(" & IIf(Val(TxtOrganizationID.Text) = 0, "Null", TxtOrganizationID.Text) & ",'" & TxtPID.Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpAdjustmentDate.DateValue + 1 & "',0),0)"
         vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpAdjustmentDate.DateValue + 1 & "',0),0)"
         vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtPID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
         LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtPID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
         
         vStrSQL = "select h.AdjustmentID, AdjustmentDate, ProductID, (SQtyPack*multiplier)+SQtyLoose as SetQty from StockAdjustmentHeader h inner join StockAdjustmentBody b on h.adjustmentid = b.adjustmentid where h.AdjustmentID < " & TxtAdjID.Text & " and ProductID = " & Val(TxtPID.Text)
         With CN.Execute(vStrSQL)
            If .RecordCount > 0 Then
               LblPre.Caption = "ID = " & !AdjustmentID & ", Date = " & Format(!AdjustmentDate, "DD/MM/yyyy") & ", SQtyLoose = " & !SetQty
            Else
               LblPre.Caption = ""
            End If
            .Close
         End With
         LblPre.Visible = True
         SubCalculateBody
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtCode.Text = ""
         TxtPID.Text = ""
         If CmbPackName.ListCount > 0 Then CmbPackName.ListIndex = 0
         TxtProductName.Text = ""
         TxtMultiplier.Text = ""
         TxtCost.Text = ""
         TxtAmount.Text = ""
         LblStock.Visible = False
         LblStockCaption.Visible = False
         LblPre.Visible = False
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         Exit Function
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
   '''''''''''''''''' ActivityLogBin For Clear Action
'      Call DeleteTempActivityLogBin(vRandomID)
      vGridRows = 0
      Grid.Redraw = False
      Grid.MoveFirst
      For vCounter = 2 To Grid.Rows
         vGridRows = vGridRows + 1
         If Trim(Grid.Columns("Code").Text) <> "" Then
            ssql = "Select Productid From StockAdjustmentbody where AdjustmentID=" & Val(TxtAdjID.Text) & " and productid = " & Val(Grid.Columns("Code").Text)
            With CN.Execute(ssql)
               If .EOF Then
                  Call ActivityLogBin("", eFrmStockAdjustment, eClearUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtAdjID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpAdjustmentDate.Date), "Cleared Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("SQtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("SQtyLoose").Text) & " Price-" & Grid.Columns("Cost").Text & " Amount-" & Grid.Columns("Amount").Text)
                  vGridRows = vGridRows - 1
               End If
            End With
         Else
            vGridRows = vGridRows - 1
         End If
         Grid.MoveNext
      Next vCounter
      If vGridRows > 0 Then Call ActivityLogBin("", eFrmStockAdjustment, eClearSavedRecord, TxtAdjID.Text, DtpAdjustmentDate.DateValue, vGridRows & " Product/s Cleared")
      Grid.Redraw = True
  ''''''''''''''''''
   FormStatus = NewMode
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

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   CN.BeginTrans
      
   Call BinData
   Call ActivityLogBin("", eFrmStockAdjustment, eDelete, TxtAdjID.Text, DtpAdjustmentDate.DateValue, Grid.Rows - 1 & " Product/s Deleted Amount: " & Val(TxtTotSAmount.Text))
   
   ''''''''''''''''''''''''''Delete Serials'''''''''''''''''''''
    If RsBodySerial.RecordCount > 0 Then
        RsBodySerial.MoveFirst
        For vCounter = 1 To RsBodySerial.RecordCount
            CN.Execute "Delete from StockAdjustmentBodySerial where AdjustmentID = " & Val(TxtAdjID.Text) & " and productid = " & Val(RsBodySerial!Productid) & " and Serial ='" & RsBodySerial!serial & "'"
            RsBodySerial.MoveNext
        Next vCounter
    End If
    
   Grid.Redraw = False
   Grid.MoveFirst
   Call ActivityLog("Stock Adjustment", eDelete, TxtAdjID.Text, DtpAdjustmentDate.DateValue)
   For vCounter = 1 To Grid.Rows
      If Trim(Grid.Columns("Code").Text) <> "" Then
         CN.Execute "Delete from StockAdjustmentBody where AdjustmentID = " & Val(TxtAdjID.Text) & " and ProductID = " & Val(Grid.Columns("Productid").Text)
      End If
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
   CN.Execute "Delete from StockAdjustmentHeader where AdjustmentID = " & Val(TxtAdjID.Text)
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnLoadData_Click()
On Error GoTo ErrorHandler
   If Val(TxtPortNumber.Text) <> 0 Then
      CipherAGB1.Port = Val(TxtPortNumber.Text)
      MsgBox IIf(CipherAGB1.InitConnection(1) = 1, "Connected", "Not Connected")
         
      CipherAGB1.ReadFile "c:\Data.txt"
   End If
   ' Get a free file number
   nFileNum = FreeFile
   
   ' Open a text file for input. inputbox returns the path to read the file
   Open "C:\Data.txt" For Input As nFileNum
   lLineCount = 1
   ' Read the contents of the file
   CN.Execute "Delete From TempData"
   
   Do While Not EOF(nFileNum)
      Line Input #nFileNum, sNextLine
      'do something with it
      'add line numbers to it, in this case!
      Dim vCommands() As String
      vCommands = Split(sNextLine, ",")
      If vCommands(0) <> "TRACK" Then
         CN.Execute "Insert Into TempData values ('" & Val(vCommands(0)) & "'," & vCommands(1) & ")"
      End If
   Loop
   ' Close the file
   Close nFileNum
   Call SubLoadData
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubLoadData()
   On Error GoTo ErrorHandler
   vStrSQL = " Select isnull(b.ProductID,t.code) as ProductID, t.Code,  sum(QtyLoose) as Qty " & vbCrLf & _
            " from TempData t left outer join ProductBarcodes b on b.Code = t.code " & vbCrLf & _
            " Group by isnull(b.ProductID,t.code), t.Code "

   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly

   If RsReport.RecordCount > 0 Then
      While Not RsReport.EOF
         TxtCode.Text = RsReport!code
         FunSelectProduct ssValidate, False
         TxtSQtyLoose.Text = RsReport!Qty
         GetDataFromTexBoxesToGrid
         DoEvents
         RsReport.MoveNext
      Wend
   End If
   
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   On Error GoTo ErrorHandler
   SchStockAdjustment.Show vbModal
   If SchStockAdjustment.ParaOutAdjustmentID <> "" Then
      TxtAdjID.Text = SchStockAdjustment.ParaOutAdjustmentID
      GetStockAdjustment
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnPrint_Click()
   On Error GoTo ErrorHandler
   BtnPrint.Enabled = False
   
   vStrSQL = " Select h.AdjustmentID, AdjustmentDate, h.StoreID, StoreName, b.ProductID, ProductName, " & vbCrLf _
      + " dbo.FunLastPurPrice(0,AdjustmentDate,b.ProductID) as Cost," & vbCrLf _
      + " isnull(multiplier,0)*isnull(OQtyPack,0) + isnull(OQtyLoose,0) as OQtyLoose, " & vbCrLf _
      + " (isnull(multiplier,0)*isnull(OQtyPack,0) + isnull(OQtyLoose,0))* dbo.FunLastPurPrice(0,AdjustmentDate,b.ProductID) as OAmount," & vbCrLf _
      + " isnull(multiplier,0)*isnull(UQtyPack,0) + isnull(UQtyLoose,0) as UQtyLoose, " & vbCrLf _
      + " (isnull(multiplier,0)*isnull(UQtyPack,0) + isnull(UQtyLoose,0))* dbo.FunLastPurPrice(0,AdjustmentDate,b.ProductID) as UAmount," & vbCrLf _
      + " isnull(multiplier,0)*isnull(SQtyPack,0) + isnull(SQtyLoose,0) as SQtyLoose, " & vbCrLf _
      + " (isnull(multiplier,0)*isnull(SQtyPack,0) + isnull(SQtyLoose,0))* dbo.FunLastPurPrice(0,AdjustmentDate,b.ProductID) as SAmount" & vbCrLf _
      + " from StockAdjustmentHeader h inner join StockAdjustmentBody b on h.Adjustmentid = b.Adjustmentid" & vbCrLf _
      + " inner join Products p on p.ProductID = b.ProductID" & vbCrLf _
      + " Left outer join Stores s on h.StoreID = s.StoreID" & vbCrLf _
      + " where h.AdjustmentID = " & Val(TxtAdjID.Text) & " Order By SerialNo"
      


   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open vStrSQL, CN, adOpenDynamic, adLockReadOnly
   Set RptReportViewer.Report = New CryRptAdjustment
      
   RptReportViewer.Report.ReportTitle = "Stock Adjustment Register"
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   RptReportViewer.Report.PaperOrientation = crPortrait
   RptReportViewer.Show
   'RptReportViewer.Report.PrintOut False
   BtnPrint.Enabled = True
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
    BtnPrint.Enabled = True
End Sub

Private Sub BtnProduct_Click()
   On Error GoTo ErrorHandler
   If FunSelectProduct(ssButton, True) = True Then
      CmbPackName.SetFocus
   Else
      TxtCode.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnProductRange_Click()
   On Error GoTo ErrorHandler
   FrmProductRangeGrid.ParaInBoth = True
   FrmProductRangeGrid.Show vbModal, Me
   RsTemp.Filter = ""
   If RsTemp.RecordCount > 0 Then
      PopulateTempToGrid
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateTempToGrid()
   On Error GoTo ErrorHandler
   With RsTemp
      .MoveFirst
      While Not .EOF
         TxtCode.Text = !Productid
         TxtPID.Text = !Productid
         FunSelectProduct ssValidate, False
         TxtSQtyPack.Text = !QtyPack
         TxtSQtyLoose.Text = !QtyLoose
         TxtSQtyLoose_Change
         GetDataFromTexBoxesToGrid
         DoEvents
         .MoveNext
      Wend
   End With
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnRefresh_Click()
   On Error GoTo ErrorHandler
         
'       Dim vBm As Variant
'    Dim lTotal As Long
'    Dim i As Integer
'
'    vBm = SSDBGrid1.Bookmark
'    SSDBGrid1.MoveFirst
'
'    For i = 0 To SSDBGrid1.Rows - 1
'        lTotal = lTotal + SSDBGrid1.Columns(1).CellValue(SSDBGrid1.GetBookmark(i))
'    Next i
'
'    Text1.Text = Format(lTotal, "currency")
'
'    SSDBGrid1.Bookmark = vBm
'
    
         
   Grid.MoveFirst
   GetDataBackFromGridToTexBoxes
   While Grid.Columns("Code").Text <> ""
'      vStrSQL = "select isnull(dbo.FunStockOrg(" & IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text) & ",'" & TxtPID.Text & "'," & TxtStoreID.Text & "," & Val(TxtPurID.Text) & "," & Val(TxtPurReturnID.Text) & "," & Val(TxtBillID.Text) & "," & Val(TxtSaleReturnID.Text) & "," & Val(TxtTransferID.Text) & "," & Val(TxtManufacturedID.Text) & ",'" & DtpTransectionDate.DateValue & "'," & Val(TxtAdjID.Text) & "),0)"
      vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpTransectionDate.DateValue & "',0),0)"
      vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
      If vIsNewRecord = False Then
         vStrSQL = "select  isnull(sum(isnull(multiplier,0)* isnull(OQtyPack,0) + OQtyLoose),0) - isnull(sum(isnull(multiplier,0)* isnull(UQtyPack,0) + UQtyLoose),0) as QtyLoose from stockadjustmentbody where adjustmentid = " & Val(TxtAdjID.Text) & " and ProductID = " & Val(TxtPID.Text)
         vQtyLoose = vQtyLoose - CN.Execute(vStrSQL).Fields(0).Value
      End If
      LblStock.Caption = vQtyLoose
      If TxtSQtyLoose.Enabled And TxtSQtyLoose.Visible Then TxtSQtyLoose.SetFocus
      Call TxtSQtyLoose_Change
      RsBody.Filter = "ProductID = " & Val(TxtPID.Text)
      With Grid
         .Columns("Code").Text = TxtCode.Text
         .Columns("ProductID").Text = TxtPID.Text
         .Columns("ProductName").Text = TxtProductName.Text
         .Columns("PackName").Text = CmbPackName.Text
         .Columns("PackingID").Text = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
         .Columns("Pack").Value = Val(TxtMultiplier.Text)
         .Columns("SQtyPack").Value = TxtSQtyPack.Text
         .Columns("SQtyLoose").Value = TxtSQtyLoose.Text
         .Columns("OQtyPack").Value = Val(TxtOQtyPack.Text) 'IIf(TxtOQtyPack.Text = "", Empty, TxtOQtyPack.Text)
         .Columns("OQtyLoose").Value = Val(TxtOQtyLoose.Text) 'IIf(TxtOQtyLoose.Text = "", Empty, TxtOQtyLoose.Text)
         .Columns("UQtyPack").Value = Val(TxtUQtyPack.Text) 'IIf(TxtUQtyPack.Text = "", Empty, TxtUQtyPack.Text)
         .Columns("UQtyLoose").Value = Val(TxtUQtyLoose.Text)   'IIf(TxtUQtyLoose.Text = "", Empty, TxtUQtyLoose.Text)
         .Columns("Cost").Value = Val(TxtCost.Text)
         .Columns("Amount").Value = Val(TxtAmount.Text)
         RsBody!code = TxtCode.Text
         RsBody!Productid = TxtPID.Text
         RsBody!PackingID = CmbPackName.ItemData(CmbPackName.ListIndex)
         RsBody!Multiplier = Val(TxtMultiplier.Text)
         RsBody!SQtyPack = Val(TxtSQtyPack.Text)
         RsBody!SQtyLoose = Val(TxtSQtyLoose.Text)
         RsBody!OQtyPack = Val(TxtOQtyPack.Text)
         RsBody!OQtyLoose = Val(TxtOQtyLoose.Text)
         RsBody!UQtyPack = Val(TxtUQtyPack.Text)
         RsBody!UQtyLoose = Val(TxtUQtyLoose.Text)
         RsBody!Cost = Val(TxtCost.Text)
         RsBody!Amount = Val(TxtAmount.Text)
      End With
      Grid.MoveNext
      GetDataBackFromGridToTexBoxes
   Wend
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
'  Header Validation
   If Trim(TxtStoreID.Text) = "" Then
      MsgBox "Enter Store ID.", vbExclamation, Me.Caption
      TxtStoreID.SetFocus
      Exit Sub
   End If
   RsBody.Filter = 0
   If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product to Adjust.", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
   End If
    ''''''''''''''''''''''''Check Organization '''''''''''''''''''''''''''''''''
  If ObjRegistry.OrganizationMandatory = True And TxtOrganizationID.Text = "" Then
    MsgBox "Please Select Organization", vbInformation, Me.Caption
    If TxtOrganizationID.Visible = True Then TxtOrganizationID.SetFocus
    Exit Sub
  End If
   If vIsNewRecord Then
      If CN.Execute("Select * from StockAdjustmentHeader where AdjustmentID = " & Val(TxtAdjID.Text)).RecordCount > 0 Then
         TxtAdjID.Text = FunGetMaxID
      End If
   End If
   'Body Validation
   ' validation has been performed when a row is added to the grid
   
   'Saving record
   CN.BeginTrans
   
   Call DeleteTempActivityLogBin(vRandomID)
   If vIsNewRecord = False Then Call ActivityLogBin("", eFrmStockAdjustment, eEdit, TxtAdjID.Text, DtpAdjustmentDate.DateValue, "Amount: " & Val(TxtTotSAmount.Text))
   
'   If vIsNewRecord = False Then Call ActivityLog("Stock Adjustment", eEdit, TxtAdjID.Text, DtpAdjustmentDate.DateValue)
   ssql = "Select * from StockAdjustmentHeader where AdjustmentID=" & Val(TxtAdjID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, CN, adOpenStatic, adLockPessimistic
      If .BOF Then
         .AddNew
         !AdjustmentID = Val(TxtAdjID.Text)
      End If
      !AdjustmentDate = DtpAdjustmentDate.DateValue
      !StoreID = TxtStoreID.Text
      !OrganizationID = IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text)
      !Remarks = IIf((TxtRemarks.Text) = "", Null, TxtRemarks.Text)
      !TotalAmount = Val(TxtDifferenceAmount.Text)
      !PurID = Val(TxtPurID.Text)
      !PurReturnID = Val(TxtPurReturnID.Text)
      !BillID = Val(TxtBillID.Text)
      !SaleReturnID = Val(TxtSaleReturnID.Text)
      !TransferID = Val(TxtTransferID.Text)
      !ManufacturedID = Val(TxtManufacturedID.Text)
      !TransectionDate = DtpTransectionDate.DateValue
      !UserNo = vUser
      .Update
      .Close
   End With
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !AdjustmentID = Val(TxtAdjID.Text)
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   
   If RsBodySerial.RecordCount > 0 Then
     With RsBodySerial
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !AdjustmentID = Val(TxtAdjID.Text)
         .MoveNext
      Next vCounter
      .UpdateBatch
     End With
   End If
   
   If vIsNewRecord = True Then Call ActivityLogBin("", eFrmStockAdjustment, eAdd, TxtAdjID.Text, DtpAdjustmentDate.DateValue, Grid.Rows - 1 & " New Product/s Added Amount: " & Val(TxtTotSAmount.Text))
'   If vIsNewRecord = True Then Call ActivityLog("Stock Adjustment", eAdd, TxtAdjID.Text, DtpAdjustmentDate.DateValue)
   CN.CommitTrans
   
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   On Error GoTo ErrorHandler
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from StockAdjustmentBody where AdjustmentID = " & Val(TxtAdjID.Text), CN, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "select p.productname, b.* from StockAdjustmentBody b join products p on p.productid = b.productid where AdjustmentID = " & Val(TxtAdjID.Text) & " Order By SerialNo"
      With CN.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         'TxtTotalAmount.Text = 0
         TxtTotSQtyPack.Text = 0
         TxtTotSQtyLoose.Text = 0
         TxtTotSAmount.Text = 0
         TxtTotOQtyPack.Text = 0
         TxtTotOQtyLoose.Text = 0
         TxtTotOAmount.Text = 0
         TxtTotUQtyPack.Text = 0
         TxtTotUQtyLoose.Text = 0
         TxtTotUAmount.Text = 0
         While Not .EOF
            Grid.AddNew
            Grid.Columns("Code").Text = !code
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("ProductName").Text = !ProductName
            If !PackingID = 0 Or IsNull(!PackingID) Then
               Grid.Columns("PackingID").Value = ""
            Else
               Grid.Columns("PackingID").Value = !PackingID
            End If
            If !PackingID = 0 Or IsNull(!PackingID) Then
               Grid.Columns("PackName").Text = ""
            Else
               Grid.Columns("PackName").Text = CN.Execute("Select PackingName from Packings where PackingID=" & !PackingID).Fields(0).Value
            End If
            Grid.Columns("Pack").Value = !Multiplier
            Grid.Columns("SQtyPack").Value = IIf(!SQtyPack = 0, "", !SQtyPack)
            Grid.Columns("SQtyLoose").Value = IIf(!SQtyLoose = 0, "", !SQtyLoose)
            Grid.Columns("OQtyPack").Value = IIf(!OQtyPack = 0, "", !OQtyPack)
            Grid.Columns("OQtyLoose").Value = IIf(!OQtyLoose = 0, "", !OQtyLoose)
            Grid.Columns("UQtyPack").Value = IIf(!UQtyPack = 0, "", !UQtyPack)
            Grid.Columns("UQtyLoose").Value = IIf(!UQtyLoose = 0, "", !UQtyLoose)
            Grid.Columns("Cost").Value = !Cost
            Grid.Columns("Amount").Value = !Amount
            Grid.Columns("IsSerial").Value = !IsSerial
            vIsSerial = !IsSerial
            TxtTotSQtyPack.Text = Val(TxtTotSQtyPack.Text) + Val(!SQtyPack)
            TxtTotSQtyLoose.Text = Val(TxtTotSQtyLoose.Text) + Val(!SQtyLoose)
            TxtTotSAmount.Text = Val(TxtTotSAmount.Text) + Val(!Amount)
            'TxtTotSAmount.Text = Val(TxtTotSAmount.Text) + Round(Val(!Cost) * (Val(!SQtyPack) * Val(!Multiplier) + Val(!SQtyLoose)), 2)
            TxtTotOQtyPack.Text = Val(TxtTotOQtyPack.Text) + Val(!OQtyPack)
            TxtTotOQtyLoose.Text = Val(TxtTotOQtyLoose.Text) + Val(!OQtyLoose)
            If Not (Val(!OQtyPack) = 0 And Val(!OQtyLoose) = 0) Then
               TxtTotOAmount.Text = Val(TxtTotOAmount.Text) + Val(!Amount)
            End If
            TxtTotUQtyPack.Text = Val(TxtTotUQtyPack.Text) + Val(!UQtyPack)
            TxtTotUQtyLoose.Text = Val(TxtTotUQtyLoose.Text) + Val(!UQtyLoose)
            If Not (Val(!UQtyLoose) = 0 And Val(!UQtyLoose) = 0) Then
               TxtTotUAmount.Text = Val(TxtTotUAmount.Text) + Val(!Amount)
            End If
            .MoveNext
         Wend
         .Close
      End With
      SubCalculateHeader
      Grid.AddNew
      Grid.Columns("ProductID").Text = " " '
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
   
   RsBodySerial.Filter = 0
   If RsBodySerial.State = adStateOpen Then RsBodySerial.Close
   RsBodySerial.Open "Select * from StockAdjustmentBodySerial where AdjustmentID =" & Val(TxtAdjID.Text) & " ", CN, adOpenDynamic, adLockBatchOptimistic
   RsBodySerial.Filter = 0
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
      Call SubClearFields
      vRandomID = Rnd() * 11111 & " " & Format(Now, "dd/mm hh:mm:ss")
      TxtAdjID.Text = FunGetMaxID()
      Call PopulateDataToGrid
      BtnPrint.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      LblStock.Visible = False
      LblStockCaption.Visible = False
      LblPre.Visible = False
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      'DtpAdjustmentDate.Enabled = True
      'If DtpAdjustmentDate.Enabled And DtpAdjustmentDate.Visible Then DtpAdjustmentDate.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      'DtpAdjustmentDate.Enabled = False
      BtnPrint.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      LblStock.Visible = False
      LblStockCaption.Visible = False
      LblPre.Visible = False
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      vIsNewRecord = False
   Case Is = ChangeMode
      BtnPrint.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub BtnStore_Click()
   On Error GoTo ErrorHandler
   If FunSelectStore(ssButton, False) = True Then
      If TxtCode.Enabled Then TxtCode.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbPackName_Click()
   On Error GoTo ErrorHandler
   If CmbPackName.Text = "" Then
      TxtMultiplier.Enabled = False
      TxtSQtyPack.Enabled = False
      TxtOQtyPack.Enabled = False
      TxtUQtyPack.Enabled = False
      TxtMultiplier.Text = ""
      TxtSQtyPack.Text = ""
      TxtOQtyPack.Text = ""
      TxtUQtyPack.Text = ""
'      TxtCost.Text = Round(vUnitPrice, 3)
   Else
      TxtSQtyPack.Enabled = True
      TxtOQtyPack.Enabled = True
      TxtUQtyPack.Enabled = True
      TxtMultiplier.Enabled = True
      If Trim(TxtCode.Text) <> "" Then
         With CN.Execute("select * from ProductPacking where productid = " & Val(TxtPID.Text) & " and PackingID=" & CmbPackName.ItemData(CmbPackName.ListIndex))
            TxtMultiplier.Text = IIf(.RecordCount = 0, "", !Multiplier)
'            If Val(TxtMultiplier.Text) <> 0 Then
'               TxtCost.Text = Round(vUnitPrice * !Multiplier, 3)
'            Else
'               TxtCost.Text = Round(vUnitPrice, 3)
'            End If
         .Close
         End With
      End If
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub DtpAdjustmentDate_Change()
   On Error GoTo ErrorHandler
   'TxtAdjID.Text = FunGetMaxID()
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   
   If KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      If TxtCode.Enabled Then
         TxtCode.SetFocus
         Call SubClearDetailArea
         RsBodySerial.Filter = ""
         RsBodySerial.Filter = "ProductID = " & Val(TxtPID.Text)
         If RsBodySerial.RecordCount > 0 Then
            RsBodySerial.Delete
            SubClearSerialFields
         End If
      End If
   ElseIf KeyCode = vbKeyF12 Then
      BtnRefresh.Visible = True
   ElseIf Shift = vbCtrlMask Then
      If ActiveControl.Name = Grid.Name Then
         If KeyCode = vbKeyDelete Then
            If Trim(Grid.Columns("Code").Text <> "") Then Call mniRemoveRow_Click
            KeyCode = 0
         Else
            KeyCode = 0: Exit Sub
         End If
      End If
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
         Case vbKeyO
            If BtnOpen.Enabled Then BtnOpen_Click
            KeyCode = 0
         Case vbKeyR
            If BtnDelete.Enabled Then BtnDelete_Click
            KeyCode = 0
'         Case vbKeyP
'            If BtnPrint.Enabled Then BtnPrint_Click
'            KeyCode = 0
         Case vbKeyL
            If BtnLoadData.Visible = False Then BtnLoadData.Visible = True
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name
         Grid_DblClick
      Case TxtCode.Name
         'If Trim(TxtCode.Text) = "" Then If BtnSave.Enabled And BtnSave.Visible Then BtnSave.SetFocus
         If FunSelectProduct(ssValidate, False) = True Then
            If ObjRegistry.SetEnterKeyGridStockAdjustment = True Then
               GetDataFromTexBoxesToGrid
            ElseIf Val(TxtMultiplier.Text) > 0 Then
               TxtSQtyPack.SetFocus
            Else
               TxtSQtyLoose.SetFocus: KeyCode = 0 'vAutoEnterBeforeQty =
            End If
         End If
      Case Else
         keybd_event 9, 1, 1, 1
               KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then If CmbPackName.Enabled Then CmbPackName.SetFocus Else TxtSQtyLoose.SetFocus
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then TxtCode.SetFocus
      End Select
   ElseIf KeyCode = vbKeyF2 Then
      If Frame1.Visible = True Then
         Frame1.Visible = False
         If TxtCode.Enabled = True Then TxtCode.SetFocus Else Grid.SetFocus
      Else
            Frame1.Visible = True
            Frame1.ZOrder 0
            If TxtSerial.Enabled = True Then TxtSerial.SetFocus
        End If
   ElseIf ActiveControl.Name = TxtCode.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtCode.Name Then
         KeyCode = 0
         If BtnSave.Enabled And BtnSave.Visible Then BtnSave.SetFocus
      End If
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrorHandler
   If KeyAscii = vbKeyReturn Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
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
   If LblHelp.FontUnderline = False Then Exit Sub
   LblHelp.FontUnderline = False
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Stock Adjustment"
   HelpLocation Me
   DtpAdjustmentDate.DateValue = Date
   With CN.Execute("Select * from Packings")
      CmbPackName.AddItem ""
      While Not .EOF
         CmbPackName.AddItem !PackingName
         CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
         .MoveNext
      Wend
      .Close
   End With
   CmbPackName.ListIndex = 0
   
   TxtStoreID.Text = IIf(IsNull(ObjRegistry.StoreID), "", ObjRegistry.StoreID)
   FunSelectStore ssValidate, True
   If CBool(ObjRegistry.StoreVisible) = True Then
      LblStoreID.Visible = True
      LblStoreName.Visible = True
      TxtStoreID.Visible = True
      TxtStoreName.Visible = True
      BtnStore.Visible = True
   Else
      LblStoreID.Visible = False
      LblStoreName.Visible = False
      TxtStoreID.Visible = False
      TxtStoreName.Visible = False
      BtnStore.Visible = False
   End If
   
   TxtOrganizationID.Text = ObjRegistry.OrganizationID
   FunSelectOrganization ssValidate, True
   TxtOrganizationID.Visible = ObjRegistry.OrganizationVisible
   BtnOrganization.Visible = ObjRegistry.OrganizationVisible
   TxtOrganizationName.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationID.Visible = ObjRegistry.OrganizationVisible
   LblOrganizationName.Visible = ObjRegistry.OrganizationVisible

   vAutoEnterBeforeQty = ObjRegistry.AutoEnterBeforeQty

   DtpTransectionDate.DateValue = Date + 1
   DtpAdjustmentDate.DateValue = Date
   FormStatus = NewMode
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   If DtpAdjustmentDate.IsDateValid = False Then Exit Function
   FunGetMaxID = CN.Execute("Select isnull(max(AdjustmentID),0)+1 from StockAdjustmentHeader").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then
            ctl.Text = ""
         End If
      End If
   Next
   'TxtTotalAmount.Text = 0
   CmbPackName.ListIndex = 0
   TxtMultiplier.Text = ""
   TxtTotSQtyPack.Text = 0
   TxtTotSQtyLoose.Text = 0
   TxtTotSAmount.Text = 0
   TxtTotOQtyPack.Text = 0
   TxtTotOQtyLoose.Text = 0
   TxtTotOAmount.Text = 0
   TxtTotUQtyPack.Text = 0
   TxtTotUQtyLoose.Text = 0
   TxtTotUAmount.Text = 0
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("ProductID").Text = " "
   Grid.Update
   Call SubClearSerialFields
   Frame1.Visible = False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Are you sure to close without save?", vbQuestion + vbApplicationModal + vbYesNo, "Alert") = vbNo Then
         Cancel = 1
      End If
   Else
    Dim frmObj As Object
    For Each frmObj In Forms
        Set frmObj = Nothing
    Next
    Set RsBody = Nothing
    Set FrmStockAdjustment = Nothing
   End If
    '''''''''''''''''' ActivityLogBin For Close Action
'      Call DeleteTempActivityLogBin(vRandomID)
      If Grid.Rows > 1 And Cancel = 0 Then
         vGridRows = 0
         Grid.Redraw = False
         Grid.MoveFirst
         For vCounter = 2 To Grid.Rows
            vGridRows = vGridRows + 1
            If Trim(Grid.Columns("Code").Text) <> "" Then
               ssql = "Select Productid From StockAdjustmentbody where AdjustmentID=" & Val(TxtAdjID.Text) & " and productid = " & Val(Grid.Columns("Code").Text)
               With CN.Execute(ssql)
                  If .EOF Then
                     Call ActivityLogBin("", eFrmStockAdjustment, eCloseUnSavedRecord, IIf(vIsNewRecord = True, "0", TxtAdjID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpAdjustmentDate.Date), "Closed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("SQtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("SQtyLoose").Text) & " Price-" & Grid.Columns("Cost").Text & " Amount-" & Grid.Columns("Amount").Text)
                     vGridRows = vGridRows - 1
                  End If
                  End With
            Else
               vGridRows = vGridRows - 1
            End If
            Grid.MoveNext
            Next vCounter
         If vGridRows > 0 Then Call ActivityLogBin("", eFrmStockAdjustment, eCloseSavedRecord, TxtAdjID.Text, DtpAdjustmentDate.DateValue, vGridRows & " Product/s Closed")
         Grid.Redraw = True
      End If
  ''''''''''''''''''
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   'TxtTotalAmount.Text = Val(TxtTotalAmount.Text) - Grid.Columns("Amount").Value
   SubCalculateHeader
   TxtTotSQtyPack.Text = Val(TxtTotSQtyPack.Text) - Val(Grid.Columns("SQtyPack").Value)
   TxtTotSQtyLoose.Text = Val(TxtTotSQtyLoose.Text) - Val(Grid.Columns("SQtyLoose").Value)
   TxtTotSAmount.Text = Val(TxtTotSAmount.Text) - Grid.Columns("Amount").Value
   TxtTotOQtyPack.Text = Val(TxtTotOQtyPack.Text) - Val(Grid.Columns("OQtyPack").Value)
   TxtTotOQtyLoose.Text = Val(TxtTotOQtyLoose.Text) - Val(Grid.Columns("OQtyLoose").Value)
   If Not (Val(TxtOQtyPack.Text) = 0 And Val(TxtOQtyLoose.Text) = 0) Then
      TxtTotOAmount.Text = Val(TxtTotOAmount.Text) - Grid.Columns("Amount").Value
   End If
   TxtTotUQtyPack.Text = Val(TxtTotUQtyPack.Text) - Val(Grid.Columns("UQtyPack").Value)
   TxtTotUQtyLoose.Text = Val(TxtTotUQtyLoose.Text) - Val(Grid.Columns("UQtyLoose").Value)
   If Not (Val(TxtUQtyPack.Text) = 0 And Val(TxtUQtyLoose.Text) = 0) Then
      TxtTotUAmount.Text = Val(TxtTotUAmount.Text) - Grid.Columns("Amount").Value
   End If
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
   On Error GoTo ErrorHandler
   Call Grid_LostFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_GotFocus()
   On Error GoTo ErrorHandler
   Flag = True
   TxtCode.Enabled = False
   BtnProduct.Enabled = False
   'TxtCode.BackColor = TxtProductName.BackColor
   'TxtCode.TabStop = False
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_LostFocus()
   On Error GoTo ErrorHandler
   Flag = False
   If Trim(Grid.Columns("ProductID").Text) = "" Then
      TxtCode.Text = ""
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      TxtCode.SetFocus
   Else
      TxtCode.Enabled = False
      BtnProduct.Enabled = False
      CmbPackName.SetFocus
      If BtnSave.Enabled = False Then BtnSave.Enabled = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("ProductID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo ErrorHandler
   If Flag Then Call GetDataBackFromGridToTexBoxes
   Call PopulateDataToGridserial
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   On Error GoTo ErrorHandler
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LblPort_Click()
'   If TxtPortNumber.Enabled Then TxtPortNumber.SetFocus
'   If TxtPortNumber.Enabled Then
'      TxtPortNumber.Enabled = False
'      LblPort.ForeColor = &H800000
'   Else
'      'TxtFirst.SetFocus
'      TxtPortNumber.Enabled = True
'      TxtPortNumber.SetFocus
'      LblPort.ForeColor = vbBlack
'   End If
'   If TxtPortNumber.Enabled = False Then
'      If CN.Execute("select * from Registry").RecordCount >= 1 Then
'         CN.Execute "update Registry set PortNumber = " & Val(TxtPortNumber.Text)
'      Else
'         CN.Execute "INSERT INTO Registry Values(" & Val(TxtPortNumber.Text) & ")"
'      End If
'   End If
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name = "Grid" Then
         If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
         ssql = "Select Productid From StockAdjustmentbody where AdjustmentID=" & Val(TxtAdjID.Text) & " and productid = " & Val(Grid.Columns("Code").Text)
         With CN.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmStockAdjustment, eRemoveRowUnSaved, IIf(vIsNewRecord = True, "0", TxtAdjID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpAdjustmentDate.Date), "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("SQtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("SQtyLoose").Text) & " Price-" & Grid.Columns("Cost").Text & " Amount-" & Grid.Columns("Amount").Text)
            Else
               Call ActivityLogBin("", eFrmStockAdjustment, eRemoveRow, TxtAdjID.Text, DtpAdjustmentDate.DateValue, "Removed Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("SQtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("SQtyLoose").Text) & " Price-" & Grid.Columns("Cost").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin(vRandomID, eFrmStockAdjustment, eAddTempRecord, TxtAdjID.Text, DtpAdjustmentDate.DateValue, "Pending Remove Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("SQtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("SQtyLoose").Text) & " Price-" & Grid.Columns("Cost").Text & " Amount-" & Grid.Columns("Amount").Text)
            End If
         End With
         
         RsBodySerial.Filter = ""
         RsBodySerial.Filter = "ProductID = " & Val(TxtCode.Text)
         If RsBodySerial.RecordCount > 0 Then RsBodySerial.Delete
          
         RsBody.Filter = "ProductID = " & Val(TxtPID.Text)
         If RsBody.RecordCount > 0 Then RsBody.Delete
         Grid.SelBookmarks.RemoveAll
         Grid.SelBookmarks.Add Grid.Bookmark
         Grid.DeleteSelected
         Grid.Refresh
         SubCalculateHeader
         RsBody.Filter = 0
         Grid.MoveLast
         GetDataBackFromGridToTexBoxes
   ElseIf Me.ActiveControl.Name = "GridSerial" Then
         If TxtCode.Enabled = True Then
           MsgBox "Please Select the parent row to delete the child row", vbInformation + vbOKOnly, "Error"
           Exit Sub
         End If
         If Trim(GridSerial.Columns("Serial").Text) = "" Then Exit Sub
         RsBodySerial.Filter = "Serial = '" & TxtSerial.Text & "'"
         If RsBodySerial.RecordCount > 0 Then RsBodySerial.Delete
         GridSerial.SelBookmarks.RemoveAll
         GridSerial.SelBookmarks.Add GridSerial.Bookmark
         GridSerial.DeleteSelected
         GridSerial.Refresh
         RsBodySerial.Filter = 0
   '    GridSerial.MoveLast
        GetDataBackFromGridSerialToTexBoxes
    End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGrid()
   On Error GoTo ErrorHandler
   Dim vrowcounter As Integer
   If Trim(TxtCode.Text) = "" Then
      TxtCode.SetFocus
      Exit Sub
   End If
   If CmbPackName.ListIndex > 0 Then
      If Trim(TxtMultiplier.Text) = 0 Then
         If TxtMultiplier.Enabled Then TxtMultiplier.SetFocus
         Exit Sub
      End If
   End If
'   If Trim(TxtOQtyPack.Text) = "" And Trim(TxtOQtyLoose.Text) = "" And Trim(TxtUQtyPack.Text) = "" And Trim(TxtUQtyLoose.Text) = "" Then
'      TxtOQtyLoose.SetFocus
'      Exit Sub
'   End If
'   With CN.Execute("select QtyLoose from CurrentStockStore where ProductID ='" & TxtPID.Text & "' and StoreID = " & Val(TxtStoreID.Text))
'      If .RecordCount > 0 Then
'         vQtyLoose = !QtyLoose
'         LblStock.Caption = !QtyLoose 'CN.Execute("SELECT isnull(dbo.FunLoseQtyToPackStr('" & TxtCode.Text & "'," & !QtyLoose & "),0)").Fields(0).Value
'      End If
'      .Close
'   End With
   RsBody.Filter = "ProductID = " & Val(TxtPID.Text)
   If TxtCode.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("Code").Text = TxtCode.Text
         RsBody!code = TxtCode.Text
         Grid.Columns("ProductID").Text = Val(TxtPID.Text)
         RsBody!Productid = TxtPID.Text
      Else
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.Rows
               If Grid.Columns("ProductID").Text = TxtPID.Text Then
               ssql = "Select Productid From StockAdjustmentbody where AdjustmentID=" & Val(TxtAdjID.Text) & " and productid = " & Val(Grid.Columns("Code").Text)
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmStockAdjustment, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtAdjID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpAdjustmentDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("SQtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("SQtyLoose").Text) & " Price-" & Grid.Columns("Cost").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmStockAdjustment, eEdit, TxtAdjID.Text, DtpAdjustmentDate.DateValue, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("SQtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("SQtyLoose").Text) & " Price-" & Grid.Columns("Cost").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  'TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  
'                  vStrSQL = "select isnull(dbo.FunStockOrg(" & IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text) & ",'" & TxtPID.Text & "'," & TxtStoreID.Text & "," & Val(TxtPurID.Text) & "," & Val(TxtPurReturnID.Text) & "," & Val(TxtBillID.Text) & "," & Val(TxtSaleReturnID.Text) & "," & Val(TxtTransferID.Text) & "," & Val(TxtManufacturedID.Text) & ",'" & DtpTransectionDate.DateValue & "'," & Val(TxtAdjID.Text) & "),0)"
                    vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpAdjustmentDate.DateValue + 1 & "',0),0)"
                  vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
                  
                  TxtSQtyLoose.Text = Val(TxtSQtyLoose.Text) + Val(Grid.Columns("SQtyLoose").Value)
                  TxtSQtyPack.Text = Val(TxtSQtyPack.Text) + Val(Grid.Columns("SQtyPack").Value)
                  TxtOQtyLoose.Text = Val(TxtOQtyLoose.Text) '+ Val(Grid.Columns("OQtyLoose").Value)
                  TxtOQtyPack.Text = Val(TxtOQtyPack.Text) '+ Val(Grid.Columns("OQtyPack").Value)
                  TxtUQtyLoose.Text = Val(TxtUQtyLoose.Text) '+ Val(Grid.Columns("UQtyLoose").Value)
                  TxtUQtyPack.Text = Val(TxtUQtyPack.Text) '+ Val(Grid.Columns("UQtyPack").Value)
                  
                  TxtTotSQtyPack.Text = Val(TxtTotSQtyPack.Text) + Val(TxtSQtyPack.Text) - Val(Grid.Columns("SQtyPack").Value)
                  TxtTotSQtyLoose.Text = Val(TxtTotSQtyLoose.Text) + Val(TxtSQtyLoose.Text) - Val(Grid.Columns("SQtyLoose").Value)
                  TxtTotOQtyPack.Text = Val(TxtTotOQtyPack.Text) + Val(TxtOQtyPack.Text) - Val(Grid.Columns("OQtyPack").Value)
                  TxtTotOQtyLoose.Text = Val(TxtTotOQtyLoose.Text) + Val(TxtOQtyLoose.Text) - Val(Grid.Columns("OQtyLoose").Value)
                  TxtTotUQtyPack.Text = Val(TxtTotUQtyPack.Text) + Val(TxtUQtyPack.Text) - Val(Grid.Columns("UQtyPack").Value)
                  TxtTotUQtyLoose.Text = Val(TxtTotUQtyLoose.Text) + Val(TxtUQtyLoose.Text) - Val(Grid.Columns("UQtyLoose").Value)
                  
                  TxtTotSAmount.Text = Val(TxtTotSAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Value)
                  
                  If (Val(TxtOQtyPack.Text) <> 0 Or Val(TxtOQtyLoose.Text) <> 0) Then
                     TxtTotOAmount.Text = Val(TxtTotOAmount.Text) + Val(TxtAmount.Text)
                  End If
                  
                  If Val(Grid.Columns("OQtyPack").Value) <> 0 Or Val(Grid.Columns("OQtyLoose").Value) <> 0 Then
                     TxtTotOAmount.Text = Val(TxtTotOAmount.Text) - Val(Grid.Columns("Amount").Text)
                  End If
                  
                  If Not (Val(TxtUQtyPack.Text) = 0 And Val(TxtUQtyLoose.Text) = 0) Then
                     TxtTotUAmount.Text = Val(TxtTotUAmount.Text) + Val(TxtAmount.Text)
                  End If

                  If Val(Grid.Columns("UQtyPack").Value) <> 0 Or Val(Grid.Columns("UQtyLoose").Value) <> 0 Then
                     TxtTotUAmount.Text = Val(TxtTotUAmount.Text) - Val(Grid.Columns("Amount").Text)
                  End If
                  
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("PackName").Text = CmbPackName.Text
                  Grid.Columns("PackingID").Text = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
                  Grid.Columns("Pack").Value = (TxtMultiplier.Text)
                  Grid.Columns("SQtyPack").Value = (TxtSQtyPack.Text)
                  Grid.Columns("SQtyLoose").Value = (TxtSQtyLoose.Text)
                  Grid.Columns("OQtyPack").Value = (TxtOQtyPack.Text)
                  Grid.Columns("OQtyLoose").Value = (TxtOQtyLoose.Text)
                  Grid.Columns("UQtyPack").Value = (TxtUQtyPack.Text)
                  Grid.Columns("UQtyLoose").Value = (TxtUQtyLoose.Text)
                  Grid.Columns("Cost").Value = (TxtCost.Text)
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  Grid.Columns("IsSerial").Value = vIsSerial
                  RsBody!PackingID = CmbPackName.ItemData(CmbPackName.ListIndex)
                  RsBody!Multiplier = Val(TxtMultiplier.Text)
                  RsBody!SQtyPack = Val(TxtSQtyPack.Text)
                  RsBody!SQtyLoose = Val(TxtSQtyLoose.Text)
                  RsBody!OQtyPack = Val(TxtOQtyPack.Text)
                  RsBody!OQtyLoose = Val(TxtOQtyLoose.Text)
                  RsBody!UQtyPack = Val(TxtUQtyPack.Text)
                  RsBody!UQtyLoose = Val(TxtUQtyLoose.Text)
                  RsBody!Cost = Val(TxtCost.Text)
                  RsBody!Amount = Val(TxtAmount.Text)
                  RsBody!IsSerial = vIsSerial
                  ssql = "Select Productid From StockAdjustmentbody where AdjustmentID=" & Val(TxtAdjID.Text) & " and productid = " & Val(Grid.Columns("Code").Text)
                  With CN.Execute(ssql)
                     If .EOF Then
                        Call ActivityLogBin("", eFrmStockAdjustment, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtAdjID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpAdjustmentDate.Date), "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("SQtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("SQtyLoose").Text) & " Price-" & Grid.Columns("Cost").Text & " Amount-" & Grid.Columns("Amount").Text)
                     Else
                        Call ActivityLogBin("", eFrmStockAdjustment, eEdit, TxtAdjID.Text, DtpAdjustmentDate.DateValue, "Updated Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("SQtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("SQtyLoose").Text) & " Price-" & Grid.Columns("Cost").Text & " Amount-" & Grid.Columns("Amount").Text)
                     End If
                  End With
                  Call ActivityLogBin(vRandomID, eFrmStockAdjustment, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtAdjID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpAdjustmentDate.Date), "Pending Update Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("SQtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("SQtyLoose").Text) & " Price-" & Grid.Columns("Cost").Text & " Amount-" & Grid.Columns("Amount").Text)
                  Grid.MoveLast
                  Call SubClearDetailArea
                  TxtCode.SetFocus
                  Grid.Redraw = True
                  Exit Sub
               End If
               Grid.MoveNext
            Next vrowcounter
         MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
'         SubClearDetailArea
'         Grid.MoveLast
         TxtCode.SetFocus
         Exit Sub
      End If
   End If
   Grid.Redraw = False
   With Grid
      If TxtCode.Enabled = True Then
         'TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text)
         TxtTotSQtyPack.Text = Val(TxtTotSQtyPack.Text) + Val(TxtSQtyPack.Text)
         TxtTotSQtyLoose.Text = Val(TxtTotSQtyLoose.Text) + Val(TxtSQtyLoose.Text)
         TxtTotSAmount.Text = Val(TxtTotSAmount.Text) + Val(TxtAmount.Text)
         TxtTotOQtyPack.Text = Val(TxtTotOQtyPack.Text) + Val(TxtOQtyPack.Text)
         TxtTotOQtyLoose.Text = Val(TxtTotOQtyLoose.Text) + Val(TxtOQtyLoose.Text)
         If Not (Val(TxtOQtyPack.Text) = 0 And Val(TxtOQtyLoose.Text) = 0) Then
            TxtTotOAmount.Text = Val(TxtTotOAmount.Text) + Val(TxtAmount.Text)
         End If
         TxtTotUQtyPack.Text = Val(TxtTotUQtyPack.Text) + Val(TxtUQtyPack.Text)
         TxtTotUQtyLoose.Text = Val(TxtTotUQtyLoose.Text) + Val(TxtUQtyLoose.Text)
         If Not (Val(TxtUQtyPack.Text) = 0 And Val(TxtUQtyLoose.Text) = 0) Then
            TxtTotUAmount.Text = Val(TxtTotUAmount.Text) + Val(TxtAmount.Text)
         End If
         If vIsNewRecord = False Then Call ActivityLogBin("", eFrmStockAdjustment, eAddNewRowByEdit, TxtAdjID.Text, DtpAdjustmentDate.DateValue, "Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtSQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtSQtyLoose.Text) & " Price-" & TxtCost.Text & " Amount-" & TxtAmount.Text)
         Call ActivityLogBin(vRandomID, eFrmStockAdjustment, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtAdjID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpAdjustmentDate.Date), "Pending Add New Code-" & TxtCode.Text & " Qty-" & Val(TxtSQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtSQtyLoose.Text) & " Price-" & TxtCost.Text & " Amount-" & TxtAmount.Text)
      Else
         'TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
         TxtTotSQtyPack.Text = Val(TxtTotSQtyPack.Text) + Val(TxtSQtyPack.Text) - Val(Grid.Columns("SQtyPack").Value)
         TxtTotSQtyLoose.Text = Val(TxtTotSQtyLoose.Text) + Val(TxtSQtyLoose.Text) - Val(Grid.Columns("SQtyLoose").Value)
         TxtTotSAmount.Text = Val(TxtTotSAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
         TxtTotOQtyPack.Text = Val(TxtTotOQtyPack.Text) + Val(TxtOQtyPack.Text) - Val(Grid.Columns("OQtyPack").Value)
         TxtTotOQtyLoose.Text = Val(TxtTotOQtyLoose.Text) + Val(TxtOQtyLoose.Text) - Val(Grid.Columns("OQtyLoose").Value)
         If Not (Val(TxtOQtyPack.Text) = 0 And Val(TxtOQtyLoose.Text) = 0) Then
            TxtTotOAmount.Text = Val(TxtTotOAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
         End If
         TxtTotUQtyPack.Text = Val(TxtTotUQtyPack.Text) + Val(TxtUQtyPack.Text) - Val(Grid.Columns("UQtyPack").Value)
         TxtTotUQtyLoose.Text = Val(TxtTotUQtyLoose.Text) + Val(TxtUQtyLoose.Text) - Val(Grid.Columns("UQtyLoose").Value)
         If Not (Val(TxtUQtyPack.Text) = 0 And Val(TxtUQtyLoose.Text) = 0) Then
            TxtTotUAmount.Text = Val(TxtTotUAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
         End If
         ssql = "Select Productid From StockAdjustmentbody where AdjustmentID=" & Val(TxtAdjID.Text) & " and productid = " & Val(Grid.Columns("Code").Text)
         With CN.Execute(ssql)
            If .EOF Then
               Call ActivityLogBin("", eFrmStockAdjustment, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtAdjID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpAdjustmentDate.Date), "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("SQtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("SQtyLoose").Text) & " Price-" & Grid.Columns("Cost").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmStockAdjustment, eEditUnSaved, IIf(vIsNewRecord = True, "0", TxtAdjID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpAdjustmentDate.Date), "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtSQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtSQtyLoose.Text) & " Price-" & TxtCost.Text & " Amount-" & TxtAmount.Text)
            Else
               Call ActivityLogBin("", eFrmStockAdjustment, eEdit, TxtAdjID.Text, DtpAdjustmentDate.Date, "Effected Code-" & Grid.Columns("Code").Text & " Qty-" & Val(Grid.Columns("SQtyPack").Text) * Val(Grid.Columns("Pack").Text) + Val(Grid.Columns("SQtyLoose").Text) & " Price-" & Grid.Columns("Cost").Text & " Amount-" & Grid.Columns("Amount").Text)
               Call ActivityLogBin("", eFrmStockAdjustment, eEdit, TxtAdjID.Text, DtpAdjustmentDate.Date, "Updated Code-" & TxtCode.Text & " Qty-" & Val(TxtSQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtSQtyLoose.Text) & " Price-" & TxtCost.Text & " Amount-" & TxtAmount.Text)
            End If
         End With
         Call ActivityLogBin(vRandomID, eFrmStockAdjustment, eAddTempRecord, IIf(vIsNewRecord = True, "0", TxtAdjID.Text), IIf(vIsNewRecord = True, "01-01-1900", DtpAdjustmentDate.Date), "Pending Update Code-" & TxtCode.Text & " Qty-" & Val(TxtSQtyPack.Text) * Val(TxtMultiplier.Text) + Val(TxtSQtyLoose.Text) & " Price-" & TxtCost.Text & " Amount-" & TxtAmount.Text)
      End If
      .Columns("Code").Text = TxtCode.Text
      .Columns("ProductID").Text = Val(TxtPID.Text)
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("PackName").Text = CmbPackName.Text
      .Columns("PackingID").Text = IIf(CmbPackName.ListIndex > 0, CmbPackName.ItemData(CmbPackName.ListIndex), "")
      .Columns("Pack").Value = Val(TxtMultiplier.Text)
      .Columns("SQtyPack").Value = TxtSQtyPack.Text
      .Columns("SQtyLoose").Value = TxtSQtyLoose.Text
      .Columns("OQtyPack").Value = Val(TxtOQtyPack.Text) 'IIf(TxtOQtyPack.Text = "", Empty, TxtOQtyPack.Text)
      .Columns("OQtyLoose").Value = Val(TxtOQtyLoose.Text) 'IIf(TxtOQtyLoose.Text = "", Empty, TxtOQtyLoose.Text)
      .Columns("UQtyPack").Value = Val(TxtUQtyPack.Text) 'IIf(TxtUQtyPack.Text = "", Empty, TxtUQtyPack.Text)
      .Columns("UQtyLoose").Value = Val(TxtUQtyLoose.Text)   'IIf(TxtUQtyLoose.Text = "", Empty, TxtUQtyLoose.Text)
      .Columns("Cost").Value = Val(TxtCost.Text)
      .Columns("Amount").Value = Val(TxtAmount.Text)
      .Columns("IsSerial").Value = vIsSerial
      RsBody!code = TxtCode.Text
      RsBody!Productid = Val(TxtPID.Text)
      RsBody!PackingID = CmbPackName.ItemData(CmbPackName.ListIndex)
      RsBody!Multiplier = Val(TxtMultiplier.Text)
      RsBody!SQtyPack = Val(TxtSQtyPack.Text)
      RsBody!SQtyLoose = Val(TxtSQtyLoose.Text)
      RsBody!OQtyPack = Val(TxtOQtyPack.Text)
      RsBody!OQtyLoose = Val(TxtOQtyLoose.Text)
      RsBody!UQtyPack = Val(TxtUQtyPack.Text)
      RsBody!UQtyLoose = Val(TxtUQtyLoose.Text)
      RsBody!Cost = Val(TxtCost.Text)
      RsBody!Amount = Val(TxtAmount.Text)
      RsBody!IsSerial = vIsSerial
      .MoveLast
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   SubCalculateHeader
   Call SubClearDetailArea
   
   TxtCode.SetFocus
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   On Error GoTo ErrorHandler
   TxtCode.Enabled = True
   BtnProduct.Enabled = True
   TxtCode.Text = ""
'   TxtPID.Text = ""
   TxtProductName.Text = ""
   CmbPackName.ListIndex = 0
   TxtMultiplier.Text = ""
   TxtSQtyPack.Text = ""
   TxtSQtyLoose.Text = ""
   TxtOQtyPack.Text = ""
   TxtOQtyLoose.Text = ""
   TxtUQtyPack.Text = ""
   TxtUQtyLoose.Text = ""
   TxtCost.Text = ""
   TxtAmount.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtCode.Text = .Columns("Code").Text
      TxtPID.Text = .Columns("ProductID").Text
      TxtProductName.Text = .Columns("ProductName").Text
      If .Columns("PackName").Text = "" Then
         CmbPackName.ListIndex = 0
      Else
         CmbPackName.Text = .Columns("PackName").Text
      End If
      TxtMultiplier.Text = .Columns("Pack").Text
      TxtSQtyLoose.Text = .Columns("SQtyLoose").Text
      TxtSQtyPack.Text = .Columns("SQtyPack").Text
      TxtOQtyLoose.Text = .Columns("OQtyLoose").Text
      TxtOQtyPack.Text = .Columns("OQtyPack").Text
      TxtUQtyLoose.Text = .Columns("UQtyLoose").Text
      TxtUQtyPack.Text = .Columns("UQtyPack").Text
      TxtCost.Text = .Columns("Cost").Text
      TxtAmount.Text = .Columns("Amount").Value
      vIsSerial = .Columns("IsSerial").Value
'      vStrSQL = "select isnull(dbo.FunStockOrg(" & IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text) & ",'" & TxtPID.Text & "'," & TxtStoreID.Text & "," & Val(TxtPurID.Text) & "," & Val(TxtPurReturnID.Text) & "," & Val(TxtBillID.Text) & "," & Val(TxtSaleReturnID.Text) & "," & Val(TxtTransferID.Text) & "," & Val(TxtManufacturedID.Text) & ",'" & DtpTransectionDate.DateValue & "'," & Val(TxtAdjID.Text) & "),0)"
        vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpAdjustmentDate.DateValue + 1 & "',0),0)"
      vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
      LblStock.Caption = vQtyLoose
      LblStock.Visible = True
      LblStockCaption.Visible = True
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
'   vStrSQL = "select isnull(dbo.FunStockOrg(" & IIf(Val(TxtOrganizationID.Text) = 0, Null, TxtOrganizationID.Text) & ",'" & TxtPID.Text & "'," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpAdjustmentDate.DateValue + 1 & "',0),0)"
        vStrSQL = "select isnull(dbo.FunStock(" & Val(TxtPID.Text) & "," & TxtStoreID.Text & ",0,0,0,0,0,0,'" & DtpAdjustmentDate.DateValue + 1 & "',0),0)"
         vQtyLoose = CN.Execute(vStrSQL).Fields(0).Value
         LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack(" & Val(TxtPID.Text) & ",Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
'         LblStock.Caption = LblStock.Caption & " " & cn.Execute("SELECT dbo.FunGetLoose('" & TxtCode.Text & "',Floor(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose(" & Val(TxtPID.Text) & ",(" & vQtyLoose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetStockAdjustment()
   On Error GoTo ErrorHandler
   ssql = "select h.*, StoreName, OrganizationName FROM StockAdjustmentHeader h inner join stores s on s.storeid = h.storeid left outer join Organizations O on O.OrganizationID = h.OrganizationID where h.AdjustmentID=" & Val(TxtAdjID.Text)
   With CN.Execute(ssql)
      If Not .BOF Then
          DtpAdjustmentDate.DateValue = !AdjustmentDate
          TxtStoreID.Text = !StoreID
          TxtStoreName.Text = !StoreName
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID) = True, "", !OrganizationID)
          TxtRemarks.Text = IIf(IsNull(!Remarks) = True, "", !Remarks)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName) = True, "", !OrganizationName)
          TxtDifferenceAmount.Text = !TotalAmount
          TxtPurID.Text = IIf(IsNull(!PurID) = True, "", !PurID)
          TxtPurReturnID.Text = IIf(IsNull(!PurReturnID) = True, "", !PurReturnID)
          TxtBillID.Text = IIf(IsNull(!BillID) = True, "", !BillID)
          TxtSaleReturnID.Text = IIf(IsNull(!SaleReturnID) = True, "", !SaleReturnID)
          TxtTransferID.Text = IIf(IsNull(!TransferID) = True, "", !TransferID)
          TxtManufacturedID.Text = IIf(IsNull(!ManufacturedID) = True, "", !ManufacturedID)
          DtpTransectionDate.DateValue = IIf(IsNull(!TransectionDate) = True, "", !TransectionDate)
      End If
      .Close
   End With
   Call PopulateDataToGrid
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtCost_Change()
   On Error GoTo ErrorHandler
   vUnitPrice = Val(TxtCost.Text)
   SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub




Private Sub TxtSerial_LostFocus()
   GetDataFromTexBoxesToGridSerial
End Sub

Private Sub TxtSQtyLoose_Change()
   On Error GoTo ErrorHandler
'   If ActiveControl.Name <> TxtSQtyLoose.Name Then Exit Sub
   If ((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose) > 0 Then
      If Val(TxtMultiplier.Text) <> 0 Then
'         TxtOQtyPack.Text = CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose) \ Val(TxtMultiplier.Text))
         TxtOQtyPack.Text = CLng((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - (vQtyLoose)) \ Val(TxtMultiplier.Text)
         '///////////// edit by farhan 18-08-2020
         TxtOQtyLoose.Text = CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - Fix(vQtyLoose)) Mod Val(TxtMultiplier.Text))
         TxtOQtyLoose.Text = Val(TxtOQtyLoose.Text) + IIf(vQtyLoose < 0, -1, 1) * (CDec(vQtyLoose) - Fix(vQtyLoose) + CDec(Val(TxtSQtyLoose.Text)) - Fix(Val(TxtSQtyLoose.Text)))
         '/////////////
      Else
         TxtOQtyLoose.Text = CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose))
      End If
      If TxtUQtyPack.Text <> "" Then TxtUQtyPack.Text = ""
      If TxtUQtyLoose.Text <> "" Then TxtUQtyLoose.Text = ""
   ElseIf ((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose) < 0 Then
      If Val(TxtMultiplier.Text) <> 0 Then
'         TxtUQtyPack.Text = Abs(CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text))) - vQtyLoose) \ Val(TxtMultiplier.Text))
         TxtUQtyPack.Text = CLng(CSng(Abs((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose)) \ Val(TxtMultiplier.Text))
         '///////////// edit by farhan 18-08-2020
         TxtUQtyLoose.Text = Abs(CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text))) - Fix(vQtyLoose)) Mod Val(TxtMultiplier.Text))
         TxtUQtyLoose.Text = Val(TxtUQtyLoose.Text) + CDec(vQtyLoose) - Fix(vQtyLoose)
         '/////////////
      Else
         TxtUQtyLoose.Text = Abs(CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose)))
      End If
      If TxtOQtyPack.Text <> "" Then TxtOQtyPack.Text = ""
      If TxtOQtyLoose.Text <> "" Then TxtOQtyLoose.Text = ""
   Else
      If TxtOQtyPack.Text <> "" Then TxtOQtyPack.Text = "0"
      If TxtOQtyLoose.Text <> "" Then TxtOQtyLoose.Text = "0"
      If TxtUQtyPack.Text <> "" Then TxtUQtyPack.Text = "0"
      If TxtUQtyLoose.Text <> "" Then TxtUQtyLoose.Text = "0"
   End If
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSQtyLoose_LostFocus()
 On Error GoTo ErrorHandler
'   Select Case ActiveControl.Name
'   Case TxtUQtyPack.Name, TxtOQtyPack.Name, TxtOQtyLoose.Name
'      Exit Sub
'   End Select
   If vIsSerial = False Then
      Call GetDataFromTexBoxesToGrid
   Else
      Frame1.Visible = True
      Frame1.ZOrder 0
      TxtSerial.Enabled = True
      TxtSerial.SetFocus
   End If
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSQtyPack_Change()
   On Error GoTo ErrorHandler
   If ((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose) > 0 Then
      If Val(TxtMultiplier.Text) <> 0 Then
'         TxtOQtyPack.Text = CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose) \ Val(TxtMultiplier.Text))
         TxtOQtyPack.Text = CLng((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - (vQtyLoose)) \ Val(TxtMultiplier.Text)
         '///////////// edit by farhan 18-08-2020
         TxtOQtyLoose.Text = CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - Fix(vQtyLoose)) Mod Val(TxtMultiplier.Text))
         TxtOQtyLoose.Text = Val(TxtOQtyLoose.Text) + IIf(vQtyLoose < 0, -1, 1) * (CDec(vQtyLoose) - Fix(vQtyLoose) + CDec(Val(TxtSQtyLoose.Text)) - Fix(Val(TxtSQtyLoose.Text)))
         '/////////////
      Else
         TxtOQtyLoose.Text = CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose))
      End If
      If TxtUQtyPack.Text <> "" Then TxtUQtyPack.Text = ""
      If TxtUQtyLoose.Text <> "" Then TxtUQtyLoose.Text = ""
   ElseIf ((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose) < 0 Then
      If Val(TxtMultiplier.Text) <> 0 Then
'         TxtUQtyPack.Text = Abs(CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text))) - vQtyLoose) \ Val(TxtMultiplier.Text))
         TxtUQtyPack.Text = CLng(CSng(Abs((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose)) \ Val(TxtMultiplier.Text))
         '///////////// edit by farhan 18-08-2020
         TxtUQtyLoose.Text = Abs(CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text))) - Fix(vQtyLoose)) Mod Val(TxtMultiplier.Text))
         TxtUQtyLoose.Text = Val(TxtUQtyLoose.Text) + CDec(vQtyLoose) - Fix(vQtyLoose)
         '/////////////
      Else
         TxtUQtyLoose.Text = Abs(CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose)))
      End If
      If TxtOQtyPack.Text <> "" Then TxtOQtyPack.Text = ""
      If TxtOQtyLoose.Text <> "" Then TxtOQtyLoose.Text = ""
   Else
      If TxtOQtyPack.Text <> "" Then TxtOQtyPack.Text = "0"
      If TxtOQtyLoose.Text <> "" Then TxtOQtyLoose.Text = "0"
      If TxtUQtyPack.Text <> "" Then TxtUQtyPack.Text = "0"
      If TxtUQtyLoose.Text <> "" Then TxtUQtyLoose.Text = "0"
   End If
   Call SubCalculateBody
'   If ActiveControl.Name <> TxtSQtyPack.Name Then Exit Sub
'   If ((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose) > 0 Then
'      If Val(TxtMultiplier.Text) <> 0 Then
''         TxtOQtyPack.Text = CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose)) \ Val(TxtMultiplier.Text)
'         TxtOQtyPack.Text = CInt(CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose)) / Val(TxtMultiplier.Text))
'         TxtOQtyLoose.Text = CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose)) Mod Val(TxtMultiplier.Text)
'      Else
'         TxtOQtyLoose.Text = CSng(((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose))
'      End If
'      If TxtUQtyPack.Text <> "" Then TxtUQtyPack.Text = ""
'      If TxtUQtyLoose.Text <> "" Then TxtUQtyLoose.Text = ""
'   ElseIf ((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose) < 0 Then
'      If Val(TxtMultiplier.Text) <> 0 Then
''         TxtUQtyPack.Text = CSng(Abs((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose)) \ Val(TxtMultiplier.Text)
'         TxtUQtyPack.Text = CInt(CSng(Abs((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose)) / Val(TxtMultiplier.Text))
'         TxtUQtyLoose.Text = CSng(Abs((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text)) - vQtyLoose)) Mod Val(TxtMultiplier.Text)
'      Else
'         TxtUQtyLoose.Text = CSng(Abs((Val(TxtMultiplier.Text) * Val(TxtSQtyPack.Text) + Val(TxtSQtyLoose.Text))) - vQtyLoose)
'      End If
'      If TxtOQtyPack.Text <> "" Then TxtOQtyPack.Text = ""
'      If TxtOQtyLoose.Text <> "" Then TxtOQtyLoose.Text = ""
'   Else
'      If TxtOQtyPack.Text <> "" Then TxtOQtyPack.Text = "0"
'      If TxtOQtyLoose.Text <> "" Then TxtOQtyLoose.Text = "0"
'      If TxtUQtyPack.Text <> "" Then TxtUQtyPack.Text = "0"
'      If TxtUQtyLoose.Text <> "" Then TxtUQtyLoose.Text = "0"
'   End If
'   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtUQtyLoose_LostFocus()
   On Error GoTo ErrorHandler
   Select Case ActiveControl.Name
   Case TxtUQtyPack.Name, TxtOQtyPack.Name, TxtOQtyLoose.Name
      Exit Sub
   End Select
   If vIsSerial = False Then
      Call GetDataFromTexBoxesToGrid
   Else
      Frame1.Visible = True
      Frame1.ZOrder 0
      TxtSerial.Enabled = True
      TxtSerial.SetFocus
   End If
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtUQtyLoose_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtUQtyLoose.Name Then Exit Sub
   If TxtOQtyPack.Text <> "" Then TxtOQtyPack.Text = ""
   If TxtOQtyLoose.Text <> "" Then TxtOQtyLoose.Text = ""
   If TxtSQtyPack.Text <> "" Then TxtSQtyPack.Text = ""
   If TxtSQtyLoose.Text <> "" Then TxtSQtyLoose.Text = ""
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtUQtyPack_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtUQtyPack.Name Then Exit Sub
   If TxtOQtyPack.Text <> "" Then TxtOQtyPack.Text = ""
   If TxtOQtyLoose.Text <> "" Then TxtOQtyLoose.Text = ""
   If TxtSQtyPack.Text <> "" Then TxtSQtyPack.Text = ""
   If TxtSQtyLoose.Text <> "" Then TxtSQtyLoose.Text = ""
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtOQtyLoose_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtOQtyLoose.Name Then Exit Sub
   If TxtUQtyPack.Text <> "" Then TxtUQtyPack.Text = ""
   If TxtUQtyLoose.Text <> "" Then TxtUQtyLoose.Text = ""
   If TxtSQtyPack.Text <> "" Then TxtSQtyPack.Text = ""
   If TxtSQtyLoose.Text <> "" Then TxtSQtyLoose.Text = ""
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtOQtyPack_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtOQtyPack.Name Then Exit Sub
   If TxtUQtyPack.Text = "" Then TxtUQtyPack.Text = ""
   If TxtUQtyLoose.Text = "" Then TxtUQtyLoose.Text = ""
   If TxtSQtyPack.Text <> "" Then TxtSQtyPack.Text = ""
   If TxtSQtyLoose.Text <> "" Then TxtSQtyLoose.Text = ""
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtMultiplier_Change()
   On Error GoTo ErrorHandler
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCode_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductName.Text = ""
      TxtCost.Text = ""
      CmbPackName.ListIndex = 0
      TxtMultiplier.Text = ""
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyDown Then Grid.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCode_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If TxtProductName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   If Trim(TxtCode.Text) = "" Then Exit Sub
   vTemp = FunSelectProduct(ssValidate, False)
   If vTemp = False Then   '   vTemp = FunSelectProduct(ssButton, False)
      Cancel = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtStoreID_Change()
   On Error GoTo ErrorHandler
   If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then TxtStoreName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtStoreID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtStoreName.Text <> "" Then Exit Sub
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
    vStrSQL = "Select * FROM Organizations where OrganizationID = " & Val(TxtOrganizationID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtOrganizationName.Text = !OrganizationName
          FunSelectOrganization = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectOrganization = False
          .Close
          TxtOrganizationID.Text = ""
          TxtOrganizationName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
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
      If TxtPurID.Enabled Then TxtPurID.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub BinData()
On Error GoTo ErrorHandler
   If ObjRegistry.UseBin = True Then
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.StockAdjustmentHeaderBin (BinDate, ActionNo, FormNo, ActionUserNo, " & TableHeaderFields(eFrmStockAdjustment) & ")" & vbCrLf _
             & "Select '" & Now & "', " & eDelete & ", " & eFrmStockAdjustment & ", " & vUser & "," & TableHeaderFields(eFrmStockAdjustment) & " from StockAdjustmentHeader " & vbCrLf _
             & "Where AdjustmentID = " & TxtAdjID.Text & " and AdjustmentDate = '" & DtpAdjustmentDate.DateValue & "'"
      CN.Execute vStrSQL
      vStrSQL = "Insert Into " & vBinDataBase & ".dbo.StockAdjustmentBodyBin (" & TableBodyFields(eFrmStockAdjustment) & ")" & vbCrLf _
             & "Select " & TableBodyFields(eFrmStockAdjustment) & " from StockAdjustmentBody " & vbCrLf _
             & "Where AdjustmentID = " & TxtAdjID.Text
      CN.Execute vStrSQL
  End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub GetDataFromTexBoxesToGridSerial()
   On Error GoTo ErrorHandler
   Dim vrowcounter As Integer
   If Trim(TxtCode.Text) = "" Then
      If TxtCode.Enabled = True Then TxtCode.SetFocus
      TxtSerial.Text = ""
      Exit Sub
   End If
     
   If Trim(TxtSerial.Text) = "" Then
      'MsgBox "Enter Product ID.", vbExclamation, "Alert"
'      TxtSerial.SetFocus
      Exit Sub
   End If
   
   
'   vStrSQL = "Select Distinct ProductID from vuPurchaseSerial where SerialAdd = 1 and Serial = '" & Trim(TxtSerial.Text) & "'"
   vStrSQL = "Select ProductID, SerialAdd from vuPurchaseSerial where Serial = '" & Trim(TxtSerial.Text) & "' Order by SerialAdd Desc"
   With CN.Execute(vStrSQL)
      If Not .EOF Then
         If !SerialAdd = True Then
            MsgBox "The Serial cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
            TxtSerial.SetFocus
            Exit Sub
          ElseIf !Productid = Val(TxtCode.Text) Then
            MsgBox "Same Serial cannot be inserted on Same Product Again", vbInformation + vbOKOnly, "Error"
            TxtSerial.SetFocus
            Exit Sub
         End If
      End If
   End With
   RsBodySerial.Filter = ""
   RsBodySerial.Filter = "Serial='" & Trim(TxtSerial.Text) & "'"
   If TxtSerial.Enabled Then
      If RsBodySerial.RecordCount = 0 Then
         RsBodySerial.AddNew
         GridSerial.Columns("ProductID").Text = TxtCode.Text
         GridSerial.Columns("Serial").Text = Trim(TxtSerial.Text)
         RsBodySerial!Productid = TxtCode.Text
         RsBodySerial!serial = Trim(TxtSerial.Text)
         RsBodySerial.Update
         TxtSerial.Text = ""
      Else
'         GridSerial.Redraw = False
'         GridSerial.MoveFirst
'            For vrowcounter = 1 To GridSerial.Rows
'               If GridSerial.Columns("Serial").Text = TxtSerial.Text Then
'                  MsgBox "The Serial cannot be inserted because it already Exist", vbInformation + vbOKOnly, "Error"
'                  'SubClearDetailArea
'                  GridSerial.MoveLast
'                  TxtSerial.SetFocus
'                  GridSerial.Redraw = True
'                  Exit Sub
'               End If
'               GridSerial.MoveNext
'            Next vrowcounter
         MsgBox "The Serial Already Exist", vbInformation + vbOKOnly, "Alert"
         
         
         
'         GridSerial.MoveLast
         RsBodySerial.Filter = "ProductID = " & Val(TxtCode.Text)
         If TxtSQtyLoose.Text = RsBodySerial.RecordCount And vIsSerial = True Then GetDataFromTexBoxesToGrid
         TxtSerial.Text = ""
         TxtSerial.SetFocus
         Exit Sub
      End If
   End If
   'GridSerial.Redraw = False
   With GridSerial
      If Trim(.Columns("Serial").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Serial").Text = " "
         .AllowAddNew = False
      End If
   End With
   
   
   RsBodySerial.Filter = "ProductID = " & Val(TxtCode.Text)
   If Val(TxtSQtyLoose.Text) = RsBodySerial.RecordCount And vIsSerial = True Then
      GetDataFromTexBoxesToGrid
      Exit Sub
   End If
   '' automove grid by enter serial No
   If ObjRegistry.AutoMoveGridWhenSerialEntered = True Then
      If Grid.Rows = Grid.Row + 1 Then Exit Sub
      Grid.MoveNext
      Call Grid_GotFocus
   End If
   GridSerial.Redraw = True
   Call GridSerial_DblClick
   If TxtSerial.Enabled = True Then TxtSerial.SetFocus
   
   
   Exit Sub
ErrorHandler:
   GridSerial.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub GridSerial_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GridSerial_DblClick()
   If Grid.Columns("Code").Text <> " " And GridSerial.Columns("Serial").Text = " " Then
        TxtSerial.Enabled = True
        TxtSerial.SetFocus
    Else
'        TxtSerial.Enabled = False
    End If
End Sub

Private Sub GridSerial_GotFocus()
   If Grid.Columns("Code").Text <> " " And GridSerial.Columns("Serial").Text = " " Then
        TxtSerial.Enabled = True
    Else
'        TxtSerial.Enabled = False
    End If
End Sub

Private Sub GridSerial_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub GridSerial_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(GridSerial.Columns("Serial").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub GridSerial_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    GetDataBackFromGridSerialToTexBoxes
End Sub

Private Sub GetDataBackFromGridSerialToTexBoxes()
   On Error GoTo ErrorHandler
   With GridSerial
      TxtSerial.Text = .Columns("Serial").Text
   End With
   If GridSerial.Rows = 1 Then GridSerial.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearSerialFields()
   On Error GoTo ErrorHandler
   TxtSerial.Text = ""
'   TxtSerial.Enabled = False
   TxtSerial.Enabled = True
   GridSerial.CancelUpdate
   GridSerial.RemoveAll
   GridSerial.AddNew
   GridSerial.Columns("Serial").Text = " "
   GridSerial.Update
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub PopulateDataToGridserial()
   If Trim(Grid.Columns("ProductID").Text) = "" Then
      RsBodySerial.Filter = 0
   Else
      RsBodySerial.Filter = 0
      RsBodySerial.Filter = "ProductID = " & Grid.Columns("ProductID").Text
   End If

   If RsBodySerial.RecordCount > 0 Then
'       sSql = "select d.* from PurchaseBodySerial d  where PurID=" & Val(TxtPurchaseID.Text) & " and Purchasedate='" & DtpPurchaseDate.DateValue & "' and ProductID = '" & Grid.Columns("ProductID").Text & "'"
'      With CN.Execute(sSql)
       With RsBodySerial
         GridSerial.Redraw = False
         GridSerial.MoveFirst
         GridSerial.RemoveAll
         GridSerial.AllowAddNew = True
         .MoveFirst
         While Not .EOF
            GridSerial.AddNew
            GridSerial.Columns("ProductID").Text = !Productid
            GridSerial.Columns("Serial").Text = !serial
            .MoveNext
         Wend
'         .Close
      End With
      GridSerial.AddNew
      GridSerial.Columns("Serial").Text = " "
      GridSerial.AllowAddNew = False
      GridSerial.Redraw = True
   Else
    Call SubClearSerialFields
   End If
   RsBodySerial.Filter = 0
End Sub

