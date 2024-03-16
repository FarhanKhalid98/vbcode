VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmMultiBarcodesDetail 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmMultiBarcodesDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbPurchasePackName 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8910
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   2100
      Width           =   1905
   End
   Begin VB.ComboBox CmbColourName 
      Height          =   315
      Left            =   4500
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2085
      Width           =   1200
   End
   Begin VB.ComboBox cmbSizeName 
      Height          =   315
      Left            =   5700
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2085
      Width           =   840
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmMultiBarcodesDetail.frx":0ECA
      Left            =   4515
      List            =   "FrmMultiBarcodesDetail.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   7425
      Width           =   3276
   End
   Begin VB.ComboBox CmbPage 
      Height          =   315
      ItemData        =   "FrmMultiBarcodesDetail.frx":0ECE
      Left            =   4515
      List            =   "FrmMultiBarcodesDetail.frx":0ED7
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6885
      Width           =   3276
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
      Height          =   4110
      Left            =   12960
      TabIndex        =   19
      Top             =   1080
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
         Height          =   3750
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         Tag             =   "NC"
         Text            =   "FrmMultiBarcodesDetail.frx":0EEE
         Top             =   360
         Width           =   3975
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
         TabIndex        =   21
         Top             =   90
         Width           =   135
      End
   End
   Begin SITextBox.Txt TxtItemCode 
      Height          =   315
      Left            =   3135
      TabIndex        =   0
      Top             =   2093
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
      IntegralPoint   =   7
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4155
      Left            =   3135
      TabIndex        =   12
      Top             =   2415
      Width           =   8790
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   10
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "FrmMultiBarcodesDetail.frx":0FBE
      AllowDelete     =   -1  'True
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedRow"
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
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   10
      Columns(0).Width=   2408
      Columns(0).Caption=   "Item Code"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   2143
      Columns(1).Caption=   "Colour"
      Columns(1).Name =   "ColourName"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1455
      Columns(2).Caption=   "Size"
      Columns(2).Name =   "SizeName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   4180
      Columns(3).Caption=   "Product Name"
      Columns(3).Name =   "Name"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   3413
      Columns(4).Caption=   "Purchase Packing Name"
      Columns(4).Name =   "PurchasePackingName"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1270
      Columns(5).Caption=   "Qty"
      Columns(5).Name =   "Qty"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).NumberFormat=   "########.##"
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "ColourID"
      Columns(6).Name =   "ColourID"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "SizeID"
      Columns(7).Name =   "SizeID"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "ProductID"
      Columns(8).Name =   "ProductID"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "PurchasePackingID"
      Columns(9).Name =   "PurchasePackingID"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      _ExtentX        =   15505
      _ExtentY        =   7329
      _StockProps     =   79
      BackColor       =   15724527
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   6420
      TabIndex        =   8
      Top             =   9360
      Width           =   1272
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Clear"
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
      MICON           =   "FrmMultiBarcodesDetail.frx":0FDA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   4140
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2078
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      TX              =   "..."
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
      MICON           =   "FrmMultiBarcodesDetail.frx":0FF6
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   6570
      TabIndex        =   14
      Top             =   2100
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   10830
      TabIndex        =   3
      Top             =   2100
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
      IntegralPoint   =   5
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Height          =   420
      Left            =   5100
      TabIndex        =   6
      Top             =   9360
      Width           =   1272
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
      MICON           =   "FrmMultiBarcodesDetail.frx":1012
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtTotQty 
      Height          =   315
      Left            =   8670
      TabIndex        =   15
      Top             =   6945
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7740
      TabIndex        =   9
      Top             =   9360
      Width           =   1272
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
      MICON           =   "FrmMultiBarcodesDetail.frx":102E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      Height          =   420
      Left            =   3780
      TabIndex        =   18
      Top             =   9360
      Width           =   1272
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Preview"
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
      MICON           =   "FrmMultiBarcodesDetail.frx":104A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtX 
      Height          =   315
      Left            =   10950
      TabIndex        =   24
      Tag             =   "NC"
      Top             =   7238
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      Alignment       =   2
      Appearance      =   0
      MaxLength       =   6
      Text            =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SITextBox.Txt TxtY 
      Height          =   315
      Left            =   11685
      TabIndex        =   25
      Tag             =   "NC"
      Top             =   7238
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      Alignment       =   2
      Appearance      =   0
      MaxLength       =   6
      Text            =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   420
      TabIndex        =   32
      Top             =   2730
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
      IntegralPoint   =   7
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Pack Name"
      Height          =   195
      Left            =   8910
      TabIndex        =   35
      Top             =   1920
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      Height          =   195
      Left            =   405
      TabIndex        =   33
      Top             =   2520
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Colour"
      Height          =   195
      Left            =   4530
      TabIndex        =   31
      Top             =   1890
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   195
      Left            =   5700
      TabIndex        =   30
      Top             =   1890
      Width           =   300
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Printer"
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
      Left            =   3840
      TabIndex        =   29
      Top             =   7470
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "H Value"
      Height          =   195
      Left            =   10950
      TabIndex        =   28
      Top             =   7028
      Width           =   570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "V Value"
      Height          =   195
      Left            =   11685
      TabIndex        =   27
      Top             =   7028
      Width           =   555
   End
   Begin VB.Label LblPrint 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "--- Print Settings ---"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   10950
      TabIndex        =   26
      Top             =   6788
      Width           =   1290
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Page Size"
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
      Left            =   3540
      TabIndex        =   23
      Top             =   6945
      Width           =   870
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
      Left            =   11295
      TabIndex        =   22
      Top             =   540
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multiple BarCodes Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   0
      Left            =   2700
      TabIndex        =   17
      Top             =   270
      Width           =   3435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity"
      Height          =   195
      Left            =   8610
      TabIndex        =   16
      Top             =   6645
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   10800
      TabIndex        =   13
      Top             =   1890
      Width           =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   6600
      TabIndex        =   11
      Top             =   1890
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code "
      Height          =   195
      Left            =   3120
      TabIndex        =   10
      Top             =   1890
      Width           =   765
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   60
      Width           =   345
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmMultiBarcodesDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DataFile As Integer, Fl As Long, Chunks As Integer
Dim Application1 As New CRAXDRT.Application
Dim Fragment As Integer, Chunk() As Byte, i As Integer, FileName As String
Const ChunkSize As Integer = 16384
Const conChunkSize = 100
Dim strFileNm As String
Dim vCounter As Integer
Dim vIsNewRecord As Boolean
Dim Rs As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim Flag As Boolean
Dim ssql As String
Dim vStrSQL As String
Dim vIsNewRow As Boolean
Dim vNoofPages As Integer
Dim vCurrentPage As Integer
Dim vRecNo As Integer
Dim vStartFrom As Integer
Dim vCurrentRecord As Integer
Dim vProductQty As Integer

Private Sub BtnClear_Click()
   SubClearFields
End Sub

Private Sub SubPrinterSetting()
   On Error GoTo ErrorHandler
   Dim vPrinter() As String
   vPrinter = Split(CmbPrinters.Text, ",")
   If cn.Execute("Select * From PrinterSetting where size = '" & CmbPage.Text & "'").RecordCount >= 1 Then
      cn.Execute "UPDATE PrinterSetting set x = " & Val(TxtX.Text) & " , y = " & Val(TxtY.Text) & ", DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "'"
   Else
      cn.Execute "INSERT INTO PrinterSetting Values(" & Val(TxtX.Text) & " ," & Val(TxtY.Text) & ",'" & CmbPage.Text & "','" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "')"
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   On Error GoTo ErrorHandler
   Call SubPrinterSetting
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub SubCalculate()
'   Dim i As Integer
'   With CN.Execute("Select * from Products where Groupid = '045'")
'      For i = 1 To .RecordCount
'         Grid.Columns("ID").Text = !ProductID
'         Grid.Columns("Name").Text = !ProductName
'         Grid.Columns("Qty").Text = "1"
'         Grid.Update
'         Grid.AddNew
'         TxtTotQty.Text = Val(TxtTotQty.Text) + 1
'         .MoveNext
'      Next i
'   End With
'End Sub

Private Sub SubBarCodeGenerate()
   On Error GoTo ErrorHandler
   If Grid.Rows = 1 Then
      MsgBox "No Product ID selected", vbOKOnly, Me.Caption
      TxtItemCode.SetFocus
      Exit Sub
   End If
   Dim vProductID  As String
   Load frmBarcode128
'   vCounter = IIf(Val(TxtStartFrom.Text) = 0, 0, Val(TxtStartFrom.Text) - 1)
'   If CmbPage.ListIndex <= 7 Then
'       vNoofPages = IIf((Val(TxtTotQty.Text) + vCounter) Mod CmbPage.ItemData(CmbPage.ListIndex) = 0, (Val(TxtTotQty.Text) + vCounter) \ CmbPage.ItemData(CmbPage.ListIndex), ((Val(TxtTotQty.Text) + vCounter) \ CmbPage.ItemData(CmbPage.ListIndex)) + 1)
'    Else
'        vNoofPages = 1
'    End If
   vCurrentRecord = 0
   cn.Execute ("delete from pic")
      
      Grid.MoveFirst
      vProductQty = Grid.Columns("Qty").Value
      vProductID = Grid.Columns("ID").Text & Right("00" + Grid.Columns("ColourID").Text, 2) & Right("00" + Grid.Columns("SizeID").Text, 2)
      frmBarcode128.TxtBarcode.Text = vProductID
      frmBarcode128.cmdDraw.Value = True
      If Rs.State = adStateOpen Then Rs.Close
      Rs.CursorLocation = adUseClient
      Rs.Open "select * from pic", cn, adOpenStatic, adLockOptimistic
      For vCounter = 1 To Val(TxtTotQty.Text)
         If vProductQty = 0 Then
            Grid.MoveNext
            vProductQty = Grid.Columns("Qty").Value
            vProductID = Grid.Columns("ID").Text & Right("00" + Grid.Columns("ColourID").Text, 2) & Right("00" + Grid.Columns("SizeID").Text, 2)
            frmBarcode128.TxtBarcode.Text = vProductID
            frmBarcode128.cmdDraw.Value = True
         End If
         Rs.AddNew
            strFileNm = "c:\" & frmBarcode128.TxtBarcode.Text & ".bmp"
            DataFile = 1
            Open strFileNm For Binary Access Read As DataFile
             Fl = LOF(DataFile)   ' Length of data in file
             If Fl = 0 Then Close DataFile: Exit Sub
             Chunks = Fl \ ChunkSize
             Fragment = Fl Mod ChunkSize
             ReDim Chunk(Fragment)
             Get DataFile, , Chunk()
             Rs(1).AppendChunk Chunk()
             ReDim Chunk(ChunkSize)
             For i = 1 To Chunks
                 Get DataFile, , Chunk()
                 Rs(1).AppendChunk Chunk()
             Next i
             Close DataFile
            vProductQty = vProductQty - 1
            'vCurrentRecord = vCurrentRecord + 1
            'vCounter = vCounter + 1
            Rs(2) = vCounter + vCurrentPage
            Rs(0) = Grid.Columns("ProductID").Text
            Rs(3) = vProductID
         Rs.Update
      Next vCounter
      vCounter = 0
      Rs.Close
      Set Rs = Nothing
      frmBarcode128.Hide
      Unload frmBarcode128
      'Call BtnPrint_Click
      'CN.Execute ("delete from pic")
'   Next vCurrentPage
   Kill "c:\*.bmp"
   'Call BtnClear_Click
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnProduct_Click()
   If FunSelectItemCode(ssButton, True) = True Then
      TxtQty.SetFocus
   Else
      TxtItemCode.SetFocus
   End If
End Sub

Private Function FunSelectItemCode(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchItemCode.Show vbModal, Me
      If SchItemCode.ParaOutItemCode = "" Then FunSelectItemCode = False: Exit Function
      TxtItemCode.Text = SchItemCode.ParaOutItemCode
   End If
   '---------------------------
   CmbPurchasePackName.Clear
   vStrSQL = "SELECT PackingID, PackingName from Packings Pk " & vbCrLf _
           + "Left outer join Products p on p.PurchasePackingID = Pk.PackingID " & vbCrLf _
           + " where ItemCode='" & TxtItemCode.Text & "'" & " and isLocked = 0 and isNoCostProduct = 0"
   With cn.Execute(vStrSQL)
      While Not .EOF
         CmbPurchasePackName.AddItem !PackingName
         CmbPurchasePackName.ItemData(CmbPurchasePackName.NewIndex) = !PackingID
         CmbPurchasePackName.ListIndex = 0
         .MoveNext
      Wend
      .Close
   End With
   vStrSQL = "SELECT ProductName, ProductID from Products where ItemCode='" & TxtItemCode.Text & "'" & " and isLocked = 0 and isNoCostProduct = 0"
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductName.Text = !ProductName
         TxtProductID.Text = !Productid
         FunSelectItemCode = True
         .Close
         Exit Function
      Else
         FunSelectItemCode = False
         .Close
         TxtItemCode.Text = ""
         TxtProductName.Text = ""
         TxtProductID.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

'Private Sub BtnProductRange_Click()
'   On Error GoTo ErrorHandler
'   SchProductRange.Show vbModal, Me
'   If SchProductRange.ParaOutFromID <> "" Then
'   Dim vPID As Long, vCounter As Long
'   vPID = SchProductRange.ParaOutFromID
'   For vCounter = CLng(SchProductRange.ParaOutFromID) To CLng(SchProductRange.ParaOutToID)
'      TxtItemCode.Text = vPID
'      FunSelectItemCode ssValidate, False
'      TxtQty.Text = SchProductRange.ParaOutQty
'      GetDataFromTexBoxesToGrid
'      vPID = vPID + 1
'      DoEvents
'   Next vCounter
'   End If
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub

Private Sub CmbPage_Click()
   On Error GoTo ErrorHandler
   With cn.Execute("select * from PrinterSetting where Size = '" & CmbPage.Text & "'")
     If .RecordCount > 0 Then
        TxtX.Text = !x
        TxtY.Text = !Y
     End If
     .Close
   End With
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub LblPrint_Click()
   TxtX.Enabled = Not TxtX.Enabled
   TxtY.Enabled = Not TxtY.Enabled
   If TxtX.Enabled Then TxtX.SetFocus
   If TxtX.Enabled Then
      LblPrint.ForeColor = vbBlack
   Else
      'TxtFirst.SetFocus
      LblPrint.ForeColor = &H800000
   End If
   If TxtX.Enabled = False Then
      Call SubPrinterSetting
   End If
End Sub

Private Sub TxtItemCode_Change()
   If ActiveControl.Name <> TxtItemCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then TxtProductName.Text = ""
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtItemCode_Validate(Cancel As Boolean)
   If TxtProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtItemCode.Text) = "" Then Exit Sub
   vTemp = FunSelectItemCode(ssValidate, False)
   If vTemp = False Then
      vTemp = FunSelectItemCode(ssButton, False)
      Cancel = False
   End If
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
   SetWindowText Me.hWnd, "Multiple Barcodes"
   
   CmbPrinters.Clear
   CmbPrinters.AddItem "Default,winspool,LPT1"
   Dim p
   For Each p In Printers
      CmbPrinters.AddItem p.DeviceName & "," & p.DriverName & "," & p.Port
   Next p
   
'   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter(0) & "' where RegistryKey = 'DeviceName'")
'   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter(1) & "' where RegistryKey = 'DriverName'")
'   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter(2) & "' where RegistryKey = 'Port'")

   CmbColourName.Clear
   With cn.Execute("Select * FROM Colours Order By ColourName")
      Do Until .EOF
         CmbColourName.AddItem !ColourName
         CmbColourName.ItemData(CmbColourName.NewIndex) = !ColourID
         .MoveNext
      Loop
   End With
     
   cmbSizeName.Clear
   With cn.Execute("Select * FROM Sizes Order By SizeName")
      Do Until .EOF
         cmbSizeName.AddItem !SizeName
         cmbSizeName.ItemData(cmbSizeName.NewIndex) = !SizeID
         .MoveNext
      Loop
   End With
   CmbColourName.ListIndex = 0
   cmbSizeName.ListIndex = 2

   CmbPage.ListIndex = 0
   SubClearFields
   With cn.Execute("select * from PrinterSetting")
     If .RecordCount > 0 Then
        TxtX.Text = !x
        TxtY.Text = !Y
        CmbPage.Text = !Size
        If Not IsNull(!DeviceName) Then
            CmbPrinters.Text = !DeviceName & "," & !DriverName & "," & !Port
        Else
            CmbPrinters.ListIndex = 0
        End If
     End If
     .Close
   End With
   TxtX.Enabled = False
   TxtY.Enabled = False
   LblPrint.ForeColor = &H800000
   HelpLocation Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name = "Grid" Then
         Grid_DblClick
      Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      If TxtItemCode.Enabled Then TxtItemCode.SetFocus: Call SubClearDetailArea
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtItemCode.Name: If FunSelectItemCode(ssFunctionKey, True) = True Then CmbColourName.SetFocus
      End Select
   ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtItemCode.Name Then
         KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      If ActiveControl.Name = Grid.Name Then
         If KeyCode = vbKeyDelete Then
            If Trim(Grid.Columns("ID").Text <> "") Then Call mniRemoveRow_Click
            KeyCode = 0
         Else
            KeyCode = 0: Exit Sub
         End If
      End If
      Select Case KeyCode
         Case vbKeyW
            If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
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
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
       RptReportViewer.Caption = "Multiple Barcode Detail"
       RptReportViewer.Show vbModal
   End If
End Sub

Private Sub BtnPrint_Click()
    If SetReport Then RptReportViewer.Report.PrintOut False
End Sub

Private Function SetReport() As Boolean
   On Error GoTo ErrorHandler
   SetReport = False
   SubBarCodeGenerate
   If RsReport.State = adStateOpen Then RsReport.Close
'   VStrSQL = "select pic.f1, pic.Barcode, p.ProductID, ProductName, ProductName1, SubGroupName, " & IIf(ChkPrice.Value = 0, IIf(ChkDiscountedPrice.Value = 0, "'' as RetailPrice", "'Rs.' + cast(cast(RetailPrice-isnull(discpc,0) as int) as varchar(10)) RetailPrice"), "'Rs.' + cast(cast(RetailPrice as int) as varchar(10)) RetailPrice") & " from pic left outer join Products p on pic.productid = p.ProductID left outer join SubGroups sg on sg.SubGroupID = p.SubGroupID Order by Sr"

   vStrSQL = "select pic.f1, pic.Barcode, p.ProductID, ProductName, packingName, 'Rs.' + cast(cast(RetailPrice as int) as varchar(10))  as RetailPrice" & vbCrLf _
   + " from pic left outer join Products p on pic.productid = p.ProductID" & vbCrLf _
   + " left outer join Packings Pk on p.purchasePackingid = pk.packingid" & vbCrLf _
   + " Order by Sr "
   RsReport.Open vStrSQL, cn, adOpenDynamic, adLockReadOnly
   If CmbPage.ListIndex = 0 Then
      Set RptReportViewer.Report = Application1.OpenReport(vTmp & "\reports\CrpMultiBarCodeContinues25X50Code128.rpt")
   End If
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyShortName 'CN.Execute("Select ShortName from company").Fields(0).Value
   Dim vPrinter() As String
   vPrinter = Split(CmbPrinters.Text, ",")
'   RptReportViewer.Report.PrinterName = vPrinter(0)
'   RptReportViewer.Report.DriverName = vPrinter(1)
'   RptReportViewer.Report.PortName = vPrinter(2)
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   RptReportViewer.Report.LeftMargin = TxtX.Text
   RptReportViewer.Report.TopMargin = TxtY.Text
   
   'RptReportViewer.Show
   'RptReportViewer.Report.PrintOut False
   'MsgBox "Print has been sent to the printer", vbInformation + vbOKOnly, "Alert"
   SetReport = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Public Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is TextBox Then
         If ctl.Tag = "" Then ctl.Text = ""
      ElseIf TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then ctl.Text = ""
      ElseIf TypeOf ctl Is ComboBox Then
      End If
   Next
   CmbPurchasePackName.Clear
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("ID").Text = " "
   Grid.Update
   BtnProduct.Enabled = True
   TxtItemCode.Enabled = True
   If TxtItemCode.Visible = True Then TxtItemCode.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGrid()
   On Error GoTo ErrorHandler
   If Trim(TxtItemCode.Text) = "" Then
      'MsgBox "Enter Group ID.", vbExclamation, "Alert"
      If TxtItemCode.Enabled = True Then TxtItemCode.SetFocus
      Exit Sub
   End If
   
   If CmbColourName.ListIndex < 0 Then
      'MsgBox "Enter Group ID.", vbExclamation, "Alert"
      CmbColourName.SetFocus
      Exit Sub
   End If
   
   If cmbSizeName.ListIndex < 0 Then
      'MsgBox "Enter Group ID.", vbExclamation, "Alert"
      cmbSizeName.SetFocus
      Exit Sub
   End If
   
   If Val(TxtQty.Text) = 0 Then
      'MsgBox "Enter Qty.", vbExclamation, "Alert"
      If TxtQty.Enabled = True Then TxtQty.SetFocus
      Exit Sub
   End If
   
'   Grid.Bookmark = vBm

   '-------------------------------------------------------------------
   If Trim(Grid.Columns("ID").Text) = "" Then
      TxtTotQty.Text = Val(TxtTotQty.Text) + Val(TxtQty.Text)
   ElseIf Trim(Grid.Columns("ID").Text) = Trim(TxtItemCode.Text) Then
      TxtTotQty.Text = Val(TxtTotQty.Text) + Val(TxtQty.Text) - Grid.Columns("Qty").Text
   Else
   
   End If
   If TxtItemCode.Enabled = True Then
         Grid.Columns("ID").Text = TxtItemCode.Text
   'Else
   '      MsgBox "The record already exist"
   '      SubClearDetailArea
   '      If TxtItemCode.Enabled Then TxtItemCode.SetFocus
   '      Exit Sub
   End If
   Grid.Redraw = False
   With Grid
      .Columns("Name").Text = TxtProductName.Text
      .Columns("ProductID").Text = TxtProductID.Text
      .Columns("ColourName").Text = CmbColourName.Text
      .Columns("ColourID").Value = IIf(CmbColourName.ListIndex > -1, CmbColourName.ItemData(CmbColourName.ListIndex), "")
      .Columns("SizeName").Text = cmbSizeName.Text
      .Columns("PurchasePackingName").Text = CmbPurchasePackName.Text
      .Columns("PurchasePackingID").Value = IIf(CmbPurchasePackName.ListIndex > 0, CmbPurchasePackName.ItemData(CmbPurchasePackName.ListIndex), "")
      .Columns("SizeID").Value = IIf(cmbSizeName.ListIndex > -1, cmbSizeName.ItemData(cmbSizeName.ListIndex), "")
      .Columns("Qty").Text = TxtQty.Text
      .MoveLast
      If Trim(.Columns("ID").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("id").Text = " "
         .AllowAddNew = False
      End If
   End With
   Call SubClearDetailArea
   TxtItemCode.SetFocus
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtQty_LostFocus()
   If Me.ActiveControl.Name = TxtItemCode.Name Then Exit Sub
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub SubClearDetailArea()
   CmbPurchasePackName.Clear
   TxtItemCode.Enabled = True
   TxtItemCode.Text = ""
   CmbColourName.ListIndex = 0
   cmbSizeName.ListIndex = 0
   TxtProductName.Text = ""
   TxtProductID.Text = ""
   TxtQty.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtItemCode.Text = .Columns("ID").Text
      TxtProductName.Text = .Columns("Name").Text
      TxtProductID.Text = .Columns("ProductID").Text
      If Trim(.Columns("ColourName").Text) = "" Then
         CmbColourName.ListIndex = 0
      Else
         CmbColourName.Text = .Columns("ColourName").Text
      End If
      If Trim(.Columns("SizeName").Text) = "" Then
         cmbSizeName.ListIndex = 0
      Else
         cmbSizeName.Text = .Columns("SizeName").Text
      End If
   CmbPurchasePackName.Clear
   vStrSQL = "SELECT PackingID, PackingName from Packings Pk " & vbCrLf _
           + "Left outer join Products p on p.PurchasePackingID = Pk.PackingID " & vbCrLf _
           + " where ItemCode='" & TxtItemCode.Text & "'" & " and isLocked = 0 and isNoCostProduct = 0"
   With cn.Execute(vStrSQL)
      While Not .EOF
         CmbPurchasePackName.AddItem !PackingName
         CmbPurchasePackName.ItemData(CmbPurchasePackName.NewIndex) = !PackingID
         CmbPurchasePackName.ListIndex = 0
         .MoveNext
      Wend
      .Close
   End With
      If Trim(.Columns("PurchasePackingName").Text) = "" Then
'         CmbPurchasePackName.ListIndex = 0
      Else
         CmbPurchasePackName.Text = .Columns("PurchasePackingName").Text
      End If
      
      TxtQty.Text = .Columns("Qty").Text
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtTotQty.Text = Val(TxtTotQty.Text) - Grid.Columns("Qty").Value
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
   Call Grid_LostFocus
End Sub

Private Sub Grid_GotFocus()
   Flag = True
   TxtItemCode.Enabled = False
  End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("ID").Text) = "" Then
      TxtItemCode.Text = ""
      TxtItemCode.Enabled = True
      TxtItemCode.SetFocus
      vIsNewRow = True
   Else
      vBm = Grid.Bookmark
      TxtItemCode.Enabled = False
      If TxtQty.Enabled = True And TxtQty.Visible Then TxtQty.SetFocus
      vIsNewRow = False
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(Grid.Columns("ID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("ID").Text) = "" Then Exit Sub
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   Grid.DeleteSelected
   Grid.SelBookmarks.RemoveAll
   Grid.Refresh
   GetDataBackFromGridToTexBoxes
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
