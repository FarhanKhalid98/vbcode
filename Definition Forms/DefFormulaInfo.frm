VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form DefFormulaInfo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "DefFormulaInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
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
      Left            =   12240
      TabIndex        =   24
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
         Height          =   3750
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   25
         Tag             =   "NC"
         Text            =   "DefFormulaInfo.frx":0ECA
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
         TabIndex        =   26
         Top             =   90
         Width           =   135
      End
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8775
      TabIndex        =   6
      Top             =   8918
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
      MICON           =   "DefFormulaInfo.frx":0FAA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7455
      TabIndex        =   3
      Top             =   8918
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
      MICON           =   "DefFormulaInfo.frx":0FC6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4815
      TabIndex        =   5
      Top             =   8918
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
      MICON           =   "DefFormulaInfo.frx":0FE2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10095
      TabIndex        =   7
      Top             =   8918
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
      MICON           =   "DefFormulaInfo.frx":0FFE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6135
      TabIndex        =   4
      Top             =   8918
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
      MICON           =   "DefFormulaInfo.frx":101A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   5685
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4508
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
      MICON           =   "DefFormulaInfo.frx":1036
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   6045
      TabIndex        =   9
      Top             =   4508
      Width           =   2580
      _ExtentX        =   4551
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   9510
      TabIndex        =   12
      Top             =   1988
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   8625
      TabIndex        =   2
      Top             =   4508
      Width           =   990
      _ExtentX        =   1746
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtFinishedProductID 
      Height          =   315
      Left            =   5415
      TabIndex        =   0
      Top             =   3158
      Width           =   1635
      _ExtentX        =   2884
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
   Begin JeweledBut.JeweledButton BtnFinishedProduct 
      Height          =   330
      Left            =   7050
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3158
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
      MICON           =   "DefFormulaInfo.frx":1052
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtFinishedProductName 
      Height          =   315
      Left            =   7410
      TabIndex        =   17
      Top             =   3158
      Width           =   2580
      _ExtentX        =   4551
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   4050
      TabIndex        =   1
      Top             =   4508
      Width           =   1635
      _ExtentX        =   2884
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   3120
      Left            =   4050
      TabIndex        =   20
      Top             =   4823
      Width           =   8100
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   7
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
      stylesets(0).Picture=   "DefFormulaInfo.frx":106E
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
      Columns.Count   =   7
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ProductID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3519
      Columns(1).Caption=   "Code"
      Columns(1).Name =   "Code"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4551
      Columns(2).Caption=   "Product Name"
      Columns(2).Name =   "ProductName"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1746
      Columns(3).Caption=   "Qty"
      Columns(3).Name =   "Qty"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
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
      Columns(5).Width=   1693
      Columns(5).Caption=   "Price"
      Columns(5).Name =   "Price"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   2275
      Columns(6).Caption=   "Amount"
      Columns(6).Name =   "Amount"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   14287
      _ExtentY        =   5503
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
   Begin SITextBox.Txt TxtID 
      Height          =   315
      Left            =   3210
      TabIndex        =   21
      Top             =   2573
      Visible         =   0   'False
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtPrice 
      Height          =   315
      Left            =   9615
      TabIndex        =   27
      Tag             =   "D"
      Top             =   4508
      Width           =   960
      _ExtentX        =   1693
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
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   10575
      TabIndex        =   29
      Tag             =   "D"
      Top             =   4508
      Width           =   1335
      _ExtentX        =   2355
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
   Begin SITextBox.Txt TxtTotalAmount 
      Height          =   315
      Left            =   10770
      TabIndex        =   31
      Top             =   8228
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Locked          =   -1  'True
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
   Begin VB.Label LblTotalAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Left            =   10800
      TabIndex        =   32
      Top             =   8003
      Width           =   945
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   10605
      TabIndex        =   30
      Top             =   4313
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   9645
      TabIndex        =   28
      Top             =   4313
      Width           =   360
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
      Left            =   11340
      TabIndex        =   23
      Top             =   630
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   195
      Left            =   3225
      TabIndex        =   22
      Top             =   2378
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Finished Product Name"
      Height          =   195
      Left            =   7410
      TabIndex        =   19
      Top             =   2963
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Finished Product ID"
      Height          =   195
      Left            =   5415
      TabIndex        =   18
      Top             =   2963
      Width           =   1395
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Formula Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   15
      Top             =   270
      Width           =   2205
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty(Loose)"
      Height          =   195
      Left            =   8610
      TabIndex        =   14
      Top             =   4313
      Width           =   765
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   9510
      TabIndex        =   13
      Top             =   1793
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   4050
      TabIndex        =   11
      Top             =   4313
      Width           =   375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   6045
      TabIndex        =   10
      Top             =   4313
      Width           =   1020
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "DefFormulaInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vCounter As Integer
Dim RsBody As New ADODB.Recordset
Dim vMaxBinID As Integer
'Dim RsReport As New ADODB.Recordset
Dim Flag As Boolean
Dim i As Integer
Dim ssql As String
Dim vStrSQL As String
'----------------------------------

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.ParaInWhere = " and isLocked = 0 and israwproduct = 1"
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtCode.Text) = "" Then Exit Function
    
    If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    vStrSQL = " SELECT p.productid, Code, ProductName " & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where ( " & IIf(IsNumeric(TxtCode.Text) = False, "", "p.productid = " & (TxtCode.Text) & " or ") & " code = '" & TxtCode.Text & "')" & " and isLocked = 0 and israwproduct = 1"
           
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtPrice.Text = cn.Execute("select dbo.FunPurPrice(" & TxtProductID.Text & ")").Fields(0).Value
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         FunSelectProduct = True
         .Close
         Exit Function
      Else
         TxtProductID.Text = ""
         TxtCode.Text = ""
         TxtProductName.Text = ""
         TxtPrice.Text = ""
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         FunSelectProduct = False
         .Close
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectFinishedProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.ParaInWhere = " and isLocked = 0 and israwproduct = 0"
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectFinishedProduct = False: Exit Function
      TxtFinishedProductID.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtFinishedProductID.Text) = "" Then Exit Function
    
    If TxtFinishedProductID.Text = "" Then FunSelectFinishedProduct = False: Exit Function
    vStrSQL = " SELECT Productid, ProductName" & vbCrLf _
           + " from Products " & vbCrLf _
           + " where productid = " & TxtFinishedProductID.Text & " and isLocked = 0 and israwproduct = 0"
  
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtFinishedProductName.Text = !ProductName
         If BtnSave.Enabled = False Then FormStatus = ChangeMode
         FunSelectFinishedProduct = True
         .Close
         Exit Function
      Else
         FunSelectFinishedProduct = False
         .Close
         TxtFinishedProductID.Text = ""
         TxtFinishedProductName.Text = ""
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
   cn.Execute ("Insert Into UserActivities values ('Formula Info'" & "," & TxtID.Text & ",Null,'Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into UserActivities values ('Formula Info'" & "," & TxtID.Text & ",Null,'Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniCashReceiveVoucher", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
'   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
'      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
'      Exit Sub
'   End If
'   If ObjUserSecurity.IsAdministrator = False Then
'    MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
'    Exit Sub
'   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   cn.BeginTrans
   
'   vMaxBinID = FunGetMaxBinID
   ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'   CN.Execute ("Insert Into Bin_ProductProcessInfoHeader Select " & vMaxBinID & ",'" & Date & "',* from ProductProcessInfoHeader Where ID = " & TxtID.Text)
    '''''''''''''''''''''''''''''''''''''''''''''''Bin Body''''''''''''''''''''''''''''''''''''''''''''''
'   CN.Execute ("Insert Into Bin_ProductProcessInfoBody Select " & vMaxBinID & ",'" & Date & "', * from ProductProcessInfoBody Where ID = " & TxtID.Text)

   '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   cn.Execute ("Insert Into UserActivities values ('Formula Info'" & "," & TxtID.Text & ",Null,'Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Grid.Redraw = False
   Grid.RemoveAll
   Call ActivityLog("Product Process Info", eDelete, TxtFinishedProductID.Text)
   cn.Execute "Delete from ProductProcessInfoBody where ID= " & (TxtID.Text)
   cn.Execute "Delete from ProductProcessInfoHeader where ID= " & (TxtID.Text)
   Grid.Redraw = True
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnFinishedProduct_Click()
    If FunSelectFinishedProduct(ssButton, True) = True Then
      TxtCode.SetFocus
   Else
      TxtFinishedProductID.SetFocus
   End If
End Sub

Private Sub BtnOpen_Click()
   On Error GoTo ErrorHandler
   SchProductProcessInfo.Show vbModal
   If SchProductProcessInfo.ParaOutID = "" Then Exit Sub
   TxtID.Text = SchProductProcessInfo.ParaOutID
   cn.Execute ("Insert Into UserActivities values ('Formula Info'" & "," & TxtID.Text & ",Null,'Opened','" & Date & "','" & Time & "',4,'Opened'," & vUser & ")")
   GetProductProcessInfo
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub BtnPrint_Click()
'On Error GoTo ErrorHandler
'   vStrSql = "select u.username, h.TransferID, h.TransferDate, h.TotalAmount as tbill, isnull(h.discount,0) as discount, isnull(h.cashReceived,0) as cashReceived, p.productname, b.qty, b.price-b.discountvalue as price, b.amount" _
'            + " from StockTransferHeader h inner join StockTransferBody b on h.TransferID = b.TransferID and h.TransferDate = b.TransferDate" _
'            + " inner join products p on p.productid = b.productid" _
'            + " inner join users u on u.UserNo = h.UserNo" _
'            + " where h.TransferID= " & Val(TxtTransferID.Text) & " and h.TransferDate='" & DtpTransferDate.DateValue & "' order by SerialNo"
'
'    If RsReport.State = adStateOpen Then RsReport.Close
'    RsReport.Open vStrSql, CN, adOpenStatic, adLockReadOnly
'
'    Set RptReportViewer.Report = New CrpPurchaseInvoice
'    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
'    RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
'    'RptReportViewer.Report.PaperSize = crPaperA4
'    'RptReportViewer.Report.PaperSize = crPaperUser
'    'RptReportViewer.Report.SetUserPaperSize 1400, 1200
'    'RptReportViewer.Report.PaperOrientation = crPortrait
'    'RptReportViewer.Show
'    RptReportViewer.Report.PrintOut False
'Exit Sub
'ErrorHandler:
'    Call ShowErrorMessage
'End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      TxtCode.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniCashReceiveVoucher", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
'  If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
'    MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
'    Exit Sub
'  End If
'  Header Validation
   If Trim(TxtFinishedProductID.Text) = "" Then
      MsgBox "Enter Finished Product ID.", vbExclamation, Me.Caption
      TxtFinishedProductID.SetFocus
      Exit Sub
   End If
   If vIsNewRecord = True Then
      If cn.Execute("select * from ProductProcessInfoHeader where FinishedProductID=" & TxtFinishedProductID.Text).RecordCount > 0 Then
         MsgBox "Finished Product ID Already Exist.", vbExclamation, Me.Caption
         TxtFinishedProductID.SetFocus
         Exit Sub
      End If
   End If
   RsBody.Filter = 0
   If RsBody.RecordCount = 0 Then
      MsgBox "Please enter at least one product", vbExclamation, "Alert"
      If TxtCode.Visible And TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
   End If
   cn.BeginTrans
   Call UserActivities
   ssql = "select * from ProductProcessInfoHeader where ID=" & Val(TxtID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open ssql, cn, adOpenStatic, adLockOptimistic
      If .BOF Then
         .AddNew
         !id = Val(TxtID.Text)
      End If
      !FinishedProductID = TxtFinishedProductID.Text
      .Update
      .Close
   End With
   If vIsNewRecord = False Then Call ActivityLog("Product Process Info", eEdit, TxtFinishedProductID.Text)
  'Body Validation
  ' validation has been performed when a row is added to the grid
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !id = TxtID.Text
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   If vIsNewRecord = True Then Call ActivityLog("Product Process Info", eAdd, TxtFinishedProductID.Text)
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub PopulateDataToGrid()
   On Error GoTo ErrorHandler
   RsBody.Filter = 0
   If RsBody.State = adStateOpen Then RsBody.Close
   RsBody.Open "Select * from ProductProcessInfoBody where ID=" & Val(TxtID.Text), cn, adOpenStatic, adLockBatchOptimistic
   If RsBody.RecordCount > 0 Then
      ssql = "select p.productname, code, b.* from ProductProcessInfoBody b join products p on p.productid = b.productid where ID=" & Val(TxtID.Text)
      With cn.Execute(ssql)
         Grid.Redraw = False
         Grid.MoveFirst
         Grid.RemoveAll
         Grid.AllowAddNew = True
         TxtTotalAmount.Text = "0"
         While Not .EOF
            Grid.AddNew
            Grid.Columns("ProductID").Text = !Productid
            Grid.Columns("Code").Text = !Code
            Grid.Columns("ProductName").Text = !ProductName
            Grid.Columns("Qty").Value = !QtyLoose
            Grid.Columns("Price").Value = Val(cn.Execute("select dbo.FunPurPrice('" & !Productid & "')").Fields(0)) * cn.Execute("select isnull(PPP.Multiplier,1) from Products p LEFT OUTER JOIN ProductPacking PPP ON PPP.ProductId = P.Productid and PPP.PackingID = p.PurchasePackingID where p.productid = " & !Productid).Fields(0).Value
            Grid.Columns("Amount").Value = Round(Val(Grid.Columns("Qty").Value) * Val(Grid.Columns("Price").Value), 3)
            TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(Grid.Columns("Amount").Value)
            .MoveNext
         Wend
         .Close
      End With
      Grid.AddNew
      Grid.Columns("Code").Text = " "
      Grid.AllowAddNew = False
      Grid.Redraw = True
   End If
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
      Call PopulateDataToGrid
      TxtID.Text = FunGetMaxID
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtFinishedProductID.Enabled = True
      BtnFinishedProduct.Enabled = True
      If TxtFinishedProductID.Enabled And TxtFinishedProductID.Visible Then TxtFinishedProductID.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      'BtnPrint.Enabled = True
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      TxtFinishedProductID.Enabled = False
      BtnFinishedProduct.Enabled = False
      vIsNewRecord = False
   Case Is = ChangeMode
      'BtnPrint.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

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
      If TxtCode.Enabled Then TxtCode.SetFocus: Call SubClearDetailArea
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
            If BtnSave.Enabled And BtnSave.Visible Then BtnSave_Click
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
            If BtnDelete.Enabled And BtnDelete.Visible Then BtnDelete_Click
            KeyCode = 0
'         Case vbKeyP
'            If BtnPrint.Enabled Then BtnPrint_Click
'            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtFinishedProductID.Name: If FunSelectFinishedProduct(ssFunctionKey, False) = True Then TxtCode.SetFocus
         Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, False) = True Then TxtCode.SetFocus
      End Select
   ElseIf ActiveControl.Name = TxtCode.Name Then
      If KeyCode = vbKeyDown Then
         Grid.SetFocus
      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtCode.Name Then
         KeyCode = 0
         If BtnSave.Enabled Then BtnSave.SetFocus
      End If
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
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
   SetWindowText Me.hWnd, "Formula Info"
   HelpLocation Me
   FormStatus = NewMode
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

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
   Grid.CancelUpdate
   Grid.RemoveAll
   Grid.AddNew
   Grid.Columns("Code").Text = " "
   Grid.Update
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
    Set DefFormulaInfo = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
   TxtTotalAmount.Text = Val(TxtTotalAmount.Text) - Grid.Columns("Amount").Value
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
   BtnProduct.Enabled = False
   'TxtCode.BackColor = TxtProductName.BackColor
   'TxtCode.TabStop = False
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete And Shift = vbShiftMask + vbCtrlMask Then mniRemoveRow_Click
End Sub

Private Sub Grid_LostFocus()
   On Error GoTo ErrorHandler
   Flag = False
   If Trim(Grid.Columns("Code").Text) = "" Then
      TxtCode.Text = ""
      TxtCode.Enabled = True
      BtnProduct.Enabled = True
      TxtCode.SetFocus
   Else
      TxtCode.Enabled = False
      BtnProduct.Enabled = False
      TxtQty.SetFocus
      If BtnSave.Enabled = False Then BtnSave.Enabled = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Trim(Grid.Columns("Code").Text) = "" Or Shift <> 0 Then Exit Sub
   If Button = 2 Then Me.PopupMenu MnuDelete
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Flag Then Call GetDataBackFromGridToTexBoxes
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub mniRemoveRow_Click()
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("Code").Text) = "" Then Exit Sub
   RsBody.Filter = "Code='" & TxtCode.Text & "'"
   If RsBody.RecordCount > 0 Then RsBody.Delete
   cn.Execute ("Insert Into UserActivities values ('Formula Info'" & "," & TxtID.Text & ",Null,'Removed ProductID-" & Grid.Columns("Code").Text & " QtyLoose- " & Grid.Columns("Qty").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   Grid.DeleteSelected
   Grid.Refresh
   RsBody.Filter = 0
   GetDataBackFromGridToTexBoxes
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetDataFromTexBoxesToGrid()
   Dim vrowcounter As Integer
   If Trim(TxtCode.Text) = "" Then
      If TxtCode.Enabled Then TxtCode.SetFocus
      Exit Sub
   End If
   If Val(TxtQty.Text) = 0 Then
      TxtQty.SetFocus
      Exit Sub
   End If
On Error GoTo ErrorHandler
   RsBody.Filter = "Code='" & TxtCode.Text & "'"
   If TxtCode.Enabled Then
      If RsBody.RecordCount = 0 Then
         RsBody.AddNew
         Grid.Columns("ProductID").Text = TxtProductID.Text
         Grid.Columns("Code").Text = TxtCode.Text
         RsBody!Productid = TxtProductID.Text
         RsBody!Code = TxtCode.Text
      Else
         Grid.Redraw = False
         Grid.MoveFirst
            For vrowcounter = 1 To Grid.Rows
               If Grid.Columns("Code").Text = TxtProductID.Text Then
                  'MsgBox "The Product cannot be inserted because it already Selected", vbInformation + vbOKOnly, "Error"
                  'SubClearDetailArea
                  TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(Grid.Columns("Amount").Text)
                  TxtQty.Text = Val(TxtQty.Text) + Grid.Columns("Qty").Value
                  Grid.Columns("ProductName").Text = TxtProductName.Text
                  Grid.Columns("Qty").Value = Val(TxtQty.Text)
                  Grid.Columns("Price").Value = Val(TxtPrice.Text)
                  Grid.Columns("Amount").Value = Val(TxtAmount.Text)
                  RsBody!QtyLoose = Val(TxtQty.Text)
                  RsBody!Price = Val(TxtPrice.Text)
                  RsBody!Amount = Val(TxtAmount.Text)
                  Grid.MoveLast
                  Call SubClearDetailArea
                  TxtCode.SetFocus
                  Grid.Redraw = True
                  Exit Sub
               End If
               Grid.MoveNext
            Next vrowcounter
         'MsgBox "The Record Already Exist", vbInformation + vbOKOnly, "Alert"
         SubClearDetailArea
         Grid.MoveLast
         TxtCode.SetFocus
         Exit Sub
      End If
   End If
   Grid.Redraw = False
   With Grid
      If TxtCode.Enabled = True Then
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text)
      Else
         TxtTotalAmount.Text = Val(TxtTotalAmount.Text) + Val(TxtAmount.Text) - Val(.Columns("Amount").Text)
      End If
      .Columns("ProductName").Text = TxtProductName.Text
      .Columns("Qty").Text = TxtQty.Text
      .Columns("Price").Value = Val(TxtPrice.Text)
      .Columns("Amount").Value = Val(TxtAmount.Text)
      RsBody!QtyLoose = Val(TxtQty.Text)
      RsBody!Price = Val(TxtPrice.Text)
      RsBody!Amount = Val(TxtAmount.Text)
      .MoveLast
      If Trim(.Columns("Code").Text) <> "" Then
         .AllowAddNew = True
         .AddNew
         .Columns("Code").Text = " "
         .AllowAddNew = False
      End If
   End With
   Call SubClearDetailArea
   TxtCode.SetFocus
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
End Sub

Private Sub SubClearDetailArea()
   TxtCode.Enabled = True
   BtnProduct.Enabled = True
   TxtCode.Text = ""
   TxtProductName.Text = ""
   TxtQty.Text = ""
   TxtPrice.Text = ""
   TxtAmount.Text = ""
End Sub

Private Sub GetDataBackFromGridToTexBoxes()
   On Error GoTo ErrorHandler
   With Grid
      TxtProductID.Text = .Columns("ProductID").Text
      TxtCode.Text = .Columns("Code").Text
      TxtProductName.Text = .Columns("ProductName").Text
      TxtQty.Text = .Columns("Qty").Text
      TxtPrice.Text = .Columns("Price").Text
      TxtAmount.Text = .Columns("Amount").Text
   End With
   If Grid.Rows = 1 Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetProductProcessInfo()
   On Error GoTo ErrorHandler
   ssql = "select FinishedProductID, ProductName FROM ProductProcessInfoHeader h join products p on p.productid = h.FinishedProductid where h.ID=" & Val(TxtID.Text)
   With cn.Execute(ssql)
      If Not .BOF Then
          TxtFinishedProductID.Text = !FinishedProductID
          TxtFinishedProductName.Text = !ProductName
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

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductName.Text = ""
   End If
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtCode_Validate(Cancel As Boolean)
   If TxtProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
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

Private Sub TxtFinishedProductID_Change()
   If TxtFinishedProductID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtFinishedProductID.Name Then Exit Sub
   If TxtFinishedProductName.Text <> "" Then
      TxtFinishedProductName.Text = ""
   End If
End Sub

Private Sub TxtFinishedProductID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtFinishedProductID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtFinishedProductName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectFinishedProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectFinishedProduct(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCalculateBody()
   On Error GoTo ErrorHandler
   TxtAmount.Text = Val(TxtQty.Text) * Val(TxtPrice.Text)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtQty_Change()
   On Error GoTo ErrorHandler
   Call SubCalculateBody
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtQty_LostFocus()
   On Error GoTo ErrorHandler
   Select Case ActiveControl.Name
   Case TxtCode.Name
      Exit Sub
   End Select
   GetDataFromTexBoxesToGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = cn.Execute("Select isnull(max(ID),0)+1 from ProductProcessInfoHeader").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxBinID = cn.Execute("Select isnull(max(BinID),0)+1 from Bin_ProductProcessInfoHeader ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False Then
      With cn.Execute("Select  * from   ProductProcessInfoHeader where ID =" & TxtID.Text)
           If TxtFinishedProductID.Text <> !FinishedProductID Then
               cn.Execute ("Insert Into UserActivities values ('Formula Info'" & "," & TxtID.Text & ", Null, 'Updated FinishedProductID-" & !FinishedProductID & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
           End If
      End With
      Grid.MoveFirst
      For i = 1 To Grid.Rows - 1
           With cn.Execute("Select * from ProductProcessInfoBody Where ID = " & TxtID.Text & " and Code ='" & Grid.Columns("Code").Text & "'")
                If .EOF = True Then
                   cn.Execute ("Insert Into UserActivities values ('Formula Info'" & "," & TxtID.Text & ",Null,'Inserted New ProductID-" & Grid.Columns("code").Text & " QtyLoose-" & Grid.Columns("Qty").Text & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                Else
                   If Grid.Columns("Qty").Text <> !QtyLoose Then
                      cn.Execute ("Insert Into UserActivities values ('Formula Info'" & "," & TxtID.Text & ",Null,'Updated ProductID-" & Grid.Columns("code").Text & " QtyLoose-" & !QtyLoose & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
                   End If
               End If
           End With
      Grid.MoveNext
      Next
   Else
      cn.Execute ("Insert Into UserActivities values ('Formula Info'" & "," & TxtID.Text & ",Null,'Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
