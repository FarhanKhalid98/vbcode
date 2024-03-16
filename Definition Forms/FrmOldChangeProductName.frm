VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmChangeProductName1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "FrmOldChangeProductName.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4043
      TabIndex        =   7
      Top             =   7950
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
      MICON           =   "FrmOldChangeProductName.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5363
      TabIndex        =   8
      Top             =   7950
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Reset"
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
      MICON           =   "FrmOldChangeProductName.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6683
      TabIndex        =   10
      Top             =   7950
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
      MICON           =   "FrmOldChangeProductName.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   1455
      Width           =   1725
      _ExtentX        =   3043
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
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   2205
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1455
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
      MICON           =   "FrmOldChangeProductName.frx":0F1E
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   2565
      TabIndex        =   1
      Top             =   1455
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   100
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5190
      Left            =   480
      TabIndex        =   9
      Top             =   2490
      Width           =   11040
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   13
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
      stylesets(0).Picture=   "FrmOldChangeProductName.frx":0F3A
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
      Columns.Count   =   13
      Columns(0).Width=   2328
      Columns(0).Caption=   "Code"
      Columns(0).Name =   "Code"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   11113
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "ProductName"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1376
      Columns(2).Caption=   "Pur Price"
      Columns(2).Name =   "PurPrice"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1455
      Columns(3).Caption=   "Retail Price"
      Columns(3).Name =   "Retailprice"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "PackingID"
      Columns(4).Name =   "PackingID"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "ProductID"
      Columns(5).Name =   "ProductID"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Caption=   "CompanyId"
      Columns(6).Name =   "CompanyId"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Caption=   "CompanyName"
      Columns(7).Name =   "CompanyName"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Caption=   "GroupId"
      Columns(8).Name =   "GroupId"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Caption=   "GroupName"
      Columns(9).Name =   "GroupName"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Caption=   "SubGroupID"
      Columns(10).Name=   "SubGroupID"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   3200
      Columns(11).Caption=   "SubGroupName"
      Columns(11).Name=   "SubGroupName"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Caption=   "BarCode"
      Columns(12).Name=   "BarCode"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   19473
      _ExtentY        =   9155
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   8010
      TabIndex        =   15
      Top             =   375
      Width           =   1050
      _ExtentX        =   1852
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
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1605
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2145
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
      MICON           =   "FrmOldChangeProductName.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   825
      TabIndex        =   4
      Top             =   2145
      Width           =   780
      _ExtentX        =   1376
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   1965
      TabIndex        =   18
      Tag             =   "nc"
      Top             =   2145
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
      Left            =   7230
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1800
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
      MICON           =   "FrmOldChangeProductName.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtSubGroupID 
      Height          =   315
      Left            =   6450
      TabIndex        =   3
      Top             =   1800
      Width           =   780
      _ExtentX        =   1376
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtSubGroupName 
      Height          =   315
      Left            =   7590
      TabIndex        =   20
      Tag             =   "nc"
      Top             =   1800
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
      Left            =   1605
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1800
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
      MICON           =   "FrmOldChangeProductName.frx":0F8E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtCompanyID 
      Height          =   315
      Left            =   825
      TabIndex        =   2
      Top             =   1800
      Width           =   780
      _ExtentX        =   1376
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtCompanyName 
      Height          =   315
      Left            =   1965
      TabIndex        =   22
      Tag             =   "nc"
      Top             =   1800
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
   Begin SITextBox.Txt TxtPurPrice 
      Height          =   315
      Left            =   8865
      TabIndex        =   5
      Top             =   1455
      Width           =   780
      _ExtentX        =   1376
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
   Begin SITextBox.Txt TxtRetailPrice 
      Height          =   315
      Left            =   9675
      TabIndex        =   6
      Top             =   1455
      Width           =   780
      _ExtentX        =   1376
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
   Begin SITextBox.Txt TxtBarCodes 
      Height          =   315
      Left            =   9345
      TabIndex        =   28
      Top             =   360
      Width           =   1905
      _ExtentX        =   3360
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bar Codes"
      Height          =   195
      Left            =   9315
      TabIndex        =   29
      Top             =   135
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Retail Price"
      Height          =   195
      Left            =   9705
      TabIndex        =   27
      Top             =   1260
      Width           =   810
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Price"
      Height          =   195
      Left            =   8940
      TabIndex        =   26
      Top             =   1260
      Width           =   645
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
      Height          =   195
      Left            =   285
      TabIndex        =   25
      Top             =   2160
      Width           =   435
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      Height          =   195
      Left            =   60
      TabIndex        =   24
      Top             =   1860
      Width           =   660
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Group"
      Height          =   195
      Left            =   5580
      TabIndex        =   23
      Top             =   1830
      Width           =   765
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   7980
      TabIndex        =   16
      Top             =   150
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   2565
      TabIndex        =   14
      Top             =   1260
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   480
      TabIndex        =   13
      Top             =   1260
      Width           =   375
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Product Name"
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
      Left            =   1920
      TabIndex        =   11
      Top             =   180
      Width           =   3855
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
Attribute VB_Name = "FrmChangeProductName1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vCounter As Integer
Dim vIsNewRecord As Boolean
Dim RsBody As New ADODB.Recordset
Dim Cn1 As New ADODB.Connection
Dim Flag As Boolean
Dim sSql As String
Dim vStrSQL As String
Dim vIsNewRow As Boolean

Private Sub BtnClear_Click()
 On Error GoTo ErrorHandler
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      TxtProductName.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

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
'    If Len(TxtCode.Text) < 5 Then
'      'TxtCode.Text = "006" + Right("0000" + CStr(Val(TxtCode.Text)), 4)
'    End If
    
    If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    vStrSQL = " SELECT p.productid, Code, ProductName, PurPrice, Retailprice,p.GroupID, GroupName, p.SubGroupID, SubGroupName, p.CompanyID, CompanyName " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " left outer join ProductPacking pp on pp.packingid = p.purchasepackingid and pp.productid = p.productid" & vbCrLf _
           + " left outer join Packings pa on pa.packingid = pp.packingid " & vbCrLf _
           + " left outer join Groups g on g.groupid = p.groupid  " & vbCrLf _
           + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID " & vbCrLf _
           + " left outer join Companies c on c.CompanyID = p.CompanyID " & vbCrLf _
           + " where p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"
  
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !ProductID
         TxtBarcodes.Text = ""
         TxtProductName.Text = !ProductName
         TxtPurPrice.Text = !PurPrice
         TxtRetailPrice.Text = !RetailPrice
         TxtGroupID.Text = IIf(IsNull(!GroupID), "", !GroupID)
         TxtGroupName.Text = IIf(IsNull(!GroupName), "", !GroupName)
         TxtSubGroupID.Text = IIf(IsNull(!SubGroupID), "", !SubGroupID)
         TxtSubGroupName.Text = IIf(IsNull(!SubGroupName), "", !SubGroupName)
         TxtCompanyID.Text = IIf(IsNull(!CompanyID), "", !CompanyID)
         TxtCompanyName.Text = IIf(IsNull(!CompanyName), "", !CompanyName)
         FunSelectProduct = True
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         TxtBarcodes.Text = IIf(Len(TxtCode.Text) > 5, TxtCode.Text, "")
         TxtCode.Text = ""
         TxtProductID.Text = FunGetMaxID
         TxtProductName.Text = ""
         TxtPurPrice.Text = ""
         TxtRetailPrice.Text = ""
         TxtGroupID.Text = ""
         TxtGroupName.Text = ""
         TxtSubGroupID.Text = ""
         TxtSubGroupName.Text = ""
         TxtCompanyID.Text = ""
         TxtCompanyName.Text = ""
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunGetMaxID() As String
   On Error GoTo ErrorHandler
   FunGetMaxID = Cn1.Execute("Select right('00000' + cast(isnull(max(cast(ProductId as smallint)),0) + 1 as varchar),5) from Products --Where ProductId like '" & TxtGroupID.Text & "%'").Fields(0)
   'FunGetMaxID = CN.Execute("Select right('0000' + cast(isnull(max(cast(substring(ProductId,3,10) as smallint)),0) + 1 as varchar),4) from Products").Fields(0) ' Where ProductId like '" & GetGroupID(CmbCompany) & "%'").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
End Sub
 
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   Else
      Dim frmObj As Object
      For Each frmObj In Forms
          Set frmObj = Nothing
      Next
      Set RsBody = Nothing
      Set FrmChangeProductName = Nothing
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductName.Text = ""
      TxtPurPrice.Text = ""
   End If
End Sub

Private Sub TxtCode_Validate(Cancel As Boolean)
   If TxtProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtCode.Text) = "" Then Exit Sub
   vTemp = FunSelectProduct(ssValidate, False)
   If vTemp = False Then
      vTemp = FunSelectProduct(ssButton, False)
      Cancel = False
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
  
  On Error GoTo ErrorHandler
  'If VIsPosted And ObjUserSecurity.IsAdministrator = False Then
  '  MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
  '  Exit Sub
  'End If
'  Header Validation
   RsBody.Filter = ""
   If Grid.Rows = 1 Then
      MsgBox "Enter atleast one product to save", vbExclamation, "Alert"
      TxtProductID.SetFocus
      Exit Sub
   End If
  'Saving record
   CN.BeginTrans
   Grid.MoveFirst
   For vCounter = 1 To Grid.Rows - 1
      If Cn1.Execute("select * from products where productid  = '" & Grid.Columns("ProductID").Text & "'").RecordCount > 0 Then
         Grid.Columns("ProductID").Text = FunGetMaxID
      End If
      sSql = "insert into products (ProductId, Productname,GroupID, SubGroupID, CompanyID, PurPrice, RetailPrice) values ('" & Grid.Columns("ProductID").Text & "','" & Replace(Grid.Columns("Productname").Text, "'", "''") & "','" & Grid.Columns("GroupID").Text & "'," & IIf(Val(Grid.Columns("SubGroupID").Value) = 0, "Null", Grid.Columns("SubGroupID").Text) & "," & IIf(Val(Grid.Columns("CompanyID").Text) = 0, "Null", Grid.Columns("CompanyID").Text) & "," & Grid.Columns("Purprice").Text & "," & Grid.Columns("Retailprice").Text & ")"
      Cn1.Execute (sSql)
      If Grid.Columns("BarCode").Text <> "" Then
         If Cn1.Execute("select * from productbarcodes where code = '" & Grid.Columns("BarCode").Text & "'").RecordCount = 0 Then
            sSql = "insert into productbarcodes (ProductId, Code) values ('" & Grid.Columns("ProductID").Text & "','" & Grid.Columns("BarCode").Text & "')"
            Cn1.Execute (sSql)
         End If
      End If

      'CN.Execute ("Update Products set ProductName = '" & Grid.Columns("ProductName").Text & "' where productid='" & Grid.Columns("ProductID").Text & "'")
      'sSql = "INSERT into ActivityLog(userno,FormType,EntryDate,Description,isnew,isedit,isdelete) values(" & ObjUserSecurity.UserNo & ",'Opening Stock', GetDate(),'ProductID = " & Grid.Columns("ProductID").Text & ", QtyLoose = " & (Val(Grid.Columns("Pack").Text) * Val(Grid.Columns("QtyPack").Value)) + Val(Grid.Columns("QtyLoose").Value) & ", Price = " & Val(Grid.Columns("PurPrice").Text) & "',1,0,0)"
      'CN.Execute sSql
      Grid.MoveNext
   Next vCounter
   'Body Validation
   ' validation has been performed when a row is added to the grid
   '-------------------------------------------------------------------------
   CN.CommitTrans
   'CN.Execute "exec SPCurrentStock"
   Grid.Redraw = True
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   'If Err.Number = -2147217873 Then Resume Next
   Grid.Redraw = True
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Change Product Name"
   If Cn1.State = adStateOpen Then Cn1.Close
   Cn1.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=NewSuper;Data Source = ."
   Cn1.CursorLocation = adUseClient
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
      'If Me.ActiveControl.Name = Grid.Name And Grid.AddItemRowIndex(Grid.Bookmark) = Grid.Rows - 1 Then
      '   Grid.Update      'End If
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtProductName.SetFocus  'CmbPackName.SetFocus
      End Select
   ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtProductID.Name Then
         KeyCode = 0
         BtnSave.SetFocus
   ElseIf Shift = vbCtrlMask Then
        Select Case KeyCode
         Case vbKeyS
             If BtnSave.Enabled Then BtnSave_Click
             KeyCode = 0
         Case vbKeyQ
             If BtnClose.Enabled Then BtnClose_Click
             KeyCode = 0
        End Select
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  If Grid.Columns(ColIndex).Text = "" Then Grid.Columns(ColIndex).Text = "0"
End Sub

Private Sub ImgExit_Click()
   Unload Me
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
      'If RsBody.State = adStateOpen Then RsBody.Close
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      'PopulateDataToGrid
      TxtCode.Enabled = True
      vIsNewRow = True
   Case Is = OpenMode
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      vIsNewRow = True
   Case Is = ChangeMode
      BtnSave.Enabled = True
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is TextBox Then
         If ctl.Tag = "" Then ctl.Text = ""
      ElseIf TypeOf ctl Is ComboBox Then
      End If
   Next
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("ProductID").Text = " "
   Grid.Update
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
'   If RsBody.State = adStateOpen Then RsBody.Close
'   RsBody.Open "Select * from OpeningStock order by productid", CN, adOpenStatic, adLockBatchOptimistic
'   If RsBody.RecordCount > 0 Then
'      '================================================
'      sSql = "select pr.Productname, os.* from OpeningStock os join Products pr on os.Productid=pr.Productid order by os.productid"
'      With CN.Execute(sSql)
'         Grid.Redraw = False
'         Grid.MoveFirst
'         Grid.RemoveAll
'         Grid.AllowAddNew = True
'         While Not .EOF
'            Grid.AddNew
'            Grid.Columns("ProductID").Text = !Productid
'            Grid.Columns("Name").Text = !ProductName
'            Grid.Columns("Qty").Value = !Qty
'            Grid.Columns("PurPrice").Value = !PurPrice
'            Grid.Columns("Amount").Value = !Amount
'            .MoveNext
'         Wend
'         .Close
'         Grid.Row = 0
'      End With
      Grid.RemoveAll
      Grid.AllowAddNew = True
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
'   End If
   Grid.FirstRow = 0
End Sub

Private Sub GetDataFromTexBoxesToGrid()
On Error GoTo ErrorHandler
   If Trim(TxtProductID.Text) = "" Then
      'MsgBox "Enter Group ID.", vbExclamation, "Alert"
      If TxtCode.Enabled = True Then TxtCode.SetFocus
      Exit Sub
   End If
      
   '-------------------------------------------------------------------

   'RsBody.Filter = "ProductID='" & TxtProductID.Text & "'"
   'If vIsNewRow Then
   '   If RsBody.RecordCount = 0 Then
   '      RsBody.AddNew
  '       RsBody!Productid = TxtProductID.Text
  '    Else
  '       MsgBox "The record already exist"
  '       SubClearDetailArea
  '       If TxtProductID.Enabled Then TxtProductID.SetFocus
  '       Exit Sub
  '    End If
  ' End If
   Grid.Redraw = False
   With Grid
      .Columns("ProductID").Text = TxtProductID.Text
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("Code").Text = TxtCode.Text
      .Columns("barCode").Text = TxtBarcodes.Text
      .Columns("PurPrice").Text = TxtPurPrice.Text
      .Columns("RetailPrice").Text = TxtRetailPrice.Text
      .Columns("GroupID").Text = TxtGroupID.Text
      .Columns("GroupName").Text = TxtGroupName.Text
      .Columns("SubGroupID").Text = TxtSubGroupID.Text
      .Columns("SubGroupName").Text = TxtSubGroupName.Text
      .Columns("CompanyID").Text = TxtCompanyID.Text
      .Columns("CompanyName").Text = TxtCompanyName.Text
      .MoveLast
      If Trim(.Columns("ProductID").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   Call SubClearDetailArea
   TxtCode.SetFocus
 '  vIsNewRow = True
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub TxtRetailPrice_LostFocus()
   'Select Case ActiveControl.Name
   'Case TxtProductName.Name, TxtCode.Name
   '   Exit Sub
   'End Select
   Call GetDataFromTexBoxesToGrid
End Sub

Private Sub SubClearDetailArea()
   TxtCode.Enabled = True
   TxtCode.Text = ""
   TxtBarcodes.Text = ""
   TxtProductName.Text = ""
   TxtPurPrice.Text = ""
   TxtRetailPrice.Text = ""
   TxtGroupID.Text = ""
   TxtGroupName.Text = ""
   TxtSubGroupID.Text = ""
   TxtSubGroupName.Text = ""
   TxtCompanyID.Text = ""
   TxtCompanyName.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtCode.Text = .Columns("Code").Text
      TxtProductID.Text = .Columns("ProductID").Text
      TxtProductName.Text = .Columns("ProductName").Text
      TxtPurPrice.Text = .Columns("PurPrice").Text
      TxtRetailPrice.Text = .Columns("RetailPrice").Text
      TxtBarcodes.Text = .Columns("BarCode").Text
      TxtGroupID.Text = .Columns("GroupID").Text
      TxtGroupName.Text = .Columns("GroupName").Text
      TxtSubGroupID.Text = .Columns("SubGroupID").Text
      TxtSubGroupName.Text = .Columns("SubGroupName").Text
      TxtCompanyID.Text = .Columns("CompanyID").Text
      TxtCompanyName.Text = .Columns("CompanyName").Text
   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
   Call Grid_LostFocus
End Sub

Private Sub Grid_GotFocus()
   Flag = True
   TxtCode.Enabled = False
End Sub

Private Sub Grid_LostFocus()
   Flag = False
   If Trim(Grid.Columns("Code").Text) = "" Then
      TxtCode.Enabled = True
      TxtCode.SetFocus
      BtnProduct.Enabled = True
      vIsNewRow = True
   Else
      TxtCode.Enabled = False
      BtnProduct.Enabled = False
      vIsNewRow = False
   End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Trim(Grid.Columns("ProductID").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   'RsBody.Filter = "ProductID='" & Grid.Columns("ProductID").Text & "'"
   'If RsBody.RecordCount > 0 Then RsBody.Delete
   'RsBody.Filter = ""
   Grid.DeleteSelected
   Grid.SelBookmarks.RemoveAll
   Grid.Refresh
   Grid.MoveLast
   GetDataBackFromGridToTexBoxes
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
