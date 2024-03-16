VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmProductRangeGrid 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmProductRangeGrid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbBrand 
      Height          =   315
      Left            =   11160
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2603
      Width           =   1665
   End
   Begin VB.ComboBox CmbSubGroup 
      Height          =   315
      Left            =   9495
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2603
      Width           =   1665
   End
   Begin VB.TextBox TxtToProductID 
      Height          =   345
      Left            =   3675
      TabIndex        =   1
      Top             =   1943
      Width           =   1380
   End
   Begin VB.ComboBox CmbSortBy 
      Height          =   315
      Left            =   11175
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1928
      Width           =   1170
   End
   Begin VB.TextBox TxtFromProductID 
      Height          =   345
      Left            =   1905
      TabIndex        =   0
      Top             =   1943
      Width           =   1380
   End
   Begin VB.TextBox TxtProductName 
      Height          =   345
      Left            =   1905
      TabIndex        =   2
      Top             =   2573
      Width           =   2700
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   5910
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2603
      Width           =   1890
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   7815
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2603
      Width           =   1665
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
      Height          =   1635
      Left            =   11790
      TabIndex        =   12
      Top             =   765
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
         Height          =   1545
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Tag             =   "NC"
         Text            =   "FrmProductRangeGrid.frx":0ECA
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
         TabIndex        =   14
         Top             =   90
         Width           =   135
      End
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   5775
      TabIndex        =   8
      Top             =   9008
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Select"
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
      MICON           =   "FrmProductRangeGrid.frx":0F1C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7080
      TabIndex        =   9
      Top             =   9008
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
      MICON           =   "FrmProductRangeGrid.frx":0F38
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8370
      TabIndex        =   10
      Top             =   9008
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
      MICON           =   "FrmProductRangeGrid.frx":0F54
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnFilter 
      Height          =   315
      Left            =   4680
      TabIndex        =   16
      Top             =   2558
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      TX              =   "Filter"
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
      MICON           =   "FrmProductRangeGrid.frx":0F70
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnApply 
      Height          =   315
      Left            =   10035
      TabIndex        =   17
      Top             =   1973
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      TX              =   "Apply All"
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
      MICON           =   "FrmProductRangeGrid.frx":0F8C
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtQtyLoose 
      Height          =   315
      Left            =   9195
      TabIndex        =   18
      Top             =   1973
      Width           =   795
      _ExtentX        =   1402
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
      Height          =   5670
      Left            =   1755
      TabIndex        =   28
      Top             =   3053
      Width           =   11850
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   8
      stylesets.count =   2
      stylesets(0).Name=   "SelectedCol"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   12713983
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
      stylesets(0).Picture=   "FrmProductRangeGrid.frx":0FA8
      stylesets(1).Name=   "SelectedRow"
      stylesets(1).ForeColor=   16777215
      stylesets(1).BackColor=   8388608
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "FrmProductRangeGrid.frx":0FC4
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedCol"
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
      SelectTypeRow   =   0
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   106
      Columns.Count   =   8
      Columns(0).Width=   1614
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   5953
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1376
      Columns(2).Caption=   "Price"
      Columns(2).Name =   "Price"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   5
      Columns(2).FieldLen=   256
      Columns(3).Width=   1085
      Columns(3).Caption=   "Qty (L)"
      Columns(3).Name =   "QtyLoose"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2461
      Columns(4).Caption=   "Company"
      Columns(4).Name =   "Company"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).NumberFormat=   "########.##"
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   2461
      Columns(5).Caption=   "Group"
      Columns(5).Name =   "Group"
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   2461
      Columns(6).Caption=   "Sub Group"
      Columns(6).Name =   "SubGroup"
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(7).Width=   2461
      Columns(7).Caption=   "Brand"
      Columns(7).Name =   "Brand"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20902
      _ExtentY        =   10001
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
   Begin JeweledBut.JeweledButton BtnFromProduct 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3315
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1943
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
      MICON           =   "FrmProductRangeGrid.frx":0FE0
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnToProduct 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5085
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1928
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
      MICON           =   "FrmProductRangeGrid.frx":0FFC
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPrice 
      Height          =   315
      Left            =   8280
      TabIndex        =   31
      Top             =   1973
      Width           =   795
      _ExtentX        =   1402
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8235
      TabIndex        =   32
      Top             =   1748
      Width           =   450
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   11160
      TabIndex        =   27
      Top             =   2378
      Width           =   1050
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubGroup Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9495
      TabIndex        =   26
      Top             =   2378
      Width           =   1395
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Product ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3675
      TabIndex        =   25
      Top             =   1703
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Loose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9150
      TabIndex        =   24
      Top             =   1748
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort BY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   11175
      TabIndex        =   23
      Top             =   1703
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Product ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1905
      TabIndex        =   22
      Top             =   1703
      Width           =   1395
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1905
      TabIndex        =   21
      Top             =   2333
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5910
      TabIndex        =   20
      Top             =   2378
      Width           =   1320
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7815
      TabIndex        =   19
      Top             =   2378
      Width           =   1065
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
      Left            =   11070
      TabIndex        =   15
      Top             =   495
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Range Grid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   11
      Top             =   270
      Width           =   2565
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "FrmProductRangeGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vWords() As String
Dim vProductName As String
Dim vProductID As String
Dim Rs As New ADODB.Recordset
Dim vSuppressUpdateEvent As Boolean
Dim vSQL As String

Private Function GetGroupID(cmb As ComboBox) As String
    On Error GoTo ErrorHandler
    If cmb.ListIndex < 0 Then Exit Function
    GetGroupID = Chr(Left(cmb.ItemData(cmb.ListIndex), 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 3, 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 5, 2))
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub PopulateGrid()
   On Error GoTo ErrorHandler
   Me.MousePointer = vbHourglass
   
   Set RsTemp = New ADODB.Recordset
   RsTemp.Fields.Append "ProductID", adInteger
   RsTemp.Fields.Append "ProductName", adVarChar, 100
   RsTemp.Fields.Append "Price", adDouble
   RsTemp.Fields.Append "QtyLoose", adDouble
   RsTemp.Open

   
   Me.MousePointer = vbHourglass
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
  
   vSQL = " SELECT ProductID, ProductName, WSPrice, GroupName, isnull(CompanyName,'') as CompanyName, isnull(SubGroupName,'') as SubGroupName, isnull(BrandName,'') as BrandName" & vbCrLf _
      + " FROM Products p Left Outer join Groups g on g.GroupID = p.GroupID" & vbCrLf _
      + " left outer join SubGroups s on s.SubGroupID = p.SubGroupID" & vbCrLf _
      + " left outer join Companies c on c.CompanyID = p.CompanyID" & vbCrLf _
      + " left outer join Brands b on b.BrandID = p.BrandID" & vbCrLf _
      + " where p.isLocked = 0 " & vProductID & vProductName & vbCrLf _
      + IIf(CmbCompany.ListIndex = 0, "", " and p.CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex)) & vbCrLf _
      + IIf(CmbGroup.ListIndex = 0, "", " and p.GroupID ='" & GetGroupID(CmbGroup) & "'") & vbCrLf _
      + IIf(CmbSubGroup.ListIndex = 0, "", " and p.SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex)) & vbCrLf _
      + IIf(CmbBrand.ListIndex = 0, "", " and p.BrandID =" & CmbBrand.ItemData(CmbBrand.ListIndex)) & vbCrLf _
      + " Order by " + CmbSortBy.Text
  
   With cn.Execute(vSQL)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("Price").Value = !WSPrice
        Grid.Columns("QtyLoose").Text = ""
        Grid.Columns("Group").Text = !GroupName
        Grid.Columns("SubGroup").Text = !SubGroupName
        Grid.Columns("Company").Text = !CompanyName
        Grid.Columns("Brand").Text = !BrandName
        Grid.Update
        .MoveNext
      Loop
   End With
   vSuppressUpdateEvent = False
   Grid.Redraw = True
   Grid.MoveFirst
   Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnApply_Click()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.rows - 1
      If Trim(TxtPrice.Text) <> "" Then
         Grid.Columns("Price").Value = Val(TxtPrice.Text)
      End If
      If Trim(TxtQtyLoose.Text) <> "" Then
         Grid.Columns("QtyLoose").Value = Val(TxtQtyLoose.Text)
      End If
      Grid.MoveNext
   Next i
   UpdateRs
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
'  Call PopulateGrid
End Sub

Private Sub BtnClose_Click()
   Set RsTemp = New ADODB.Recordset
   RsTemp.Fields.Append "ProductID", adInteger
   RsTemp.Fields.Append "ProductName", adVarChar, 100
   RsTemp.Fields.Append "Price", adDouble
   RsTemp.Fields.Append "QtyLoose", adDouble
   RsTemp.Open
   Unload Me
End Sub

Private Sub BtnFilter_Click()
   On Error GoTo ErrorHandler
   vProductID = IIf(Val(TxtToProductID.Text) = 0, IIf(Val(TxtFromProductID.Text) = 0, "", " and ProductID = " & Val(TxtFromProductID.Text)), IIf(Val(TxtFromProductID.Text) = 0, "", " and ProductID Between " & Val(TxtFromProductID.Text) & " and " & Val(TxtToProductID.Text)))
   vWords = Split(TxtProductName.Text, " ")
   vProductName = ""
   For i = 0 To UBound(vWords)
       vProductName = vProductName & " and Productname like '%" & Replace(vWords(i), "'", "''") & "%'"
   Next
   PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnFromProduct_Click()
   If FunSelectFromProduct(ssButton, True) = True Then
      TxtToProductID.SetFocus
   Else
      TxtFromProductID.SetFocus
   End If
End Sub

Private Sub BtnSelect_Click()
   On Error GoTo ErrorHandler
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnToProduct_Click()
   If FunSelectToProduct(ssButton, True) = True Then
      TxtProductName.SetFocus
   Else
      TxtToProductID.SetFocus
   End If
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
   SetWindowText Me.hWnd, "Product Range Grid"
   HelpLocation Me
 
   CmbSortBy.Clear
   CmbSortBy.AddItem "ProductID"
   CmbSortBy.AddItem "ProductName"
   
   CmbBrand.Clear
   With cn.Execute("Select * FROM Brands Order By BrandName")
      CmbBrand.AddItem "All Brands"
      CmbBrand.ItemData(CmbBrand.NewIndex) = 0
      Do Until .EOF
         CmbBrand.AddItem !BrandName
         CmbBrand.ItemData(CmbBrand.NewIndex) = !BrandID
         .MoveNext
      Loop
   End With
   
   CmbCompany.Clear
   With cn.Execute("Select * FROM Companies Order By CompanyName")
      CmbCompany.AddItem "All Companies"
      CmbCompany.ItemData(CmbCompany.NewIndex) = 0
      Do Until .EOF
         CmbCompany.AddItem !CompanyName
         CmbCompany.ItemData(CmbCompany.NewIndex) = !companyid
         .MoveNext
      Loop
   End With
   
   CmbGroup.Clear
   With cn.Execute("Select * FROM Groups Order By GroupName")
      CmbGroup.AddItem "All Groups"
      CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left("000", 1)) & Asc(Mid("000", 2, 1)) & Asc(Mid("000", 3, 1))
      Do Until .EOF
         CmbGroup.AddItem !GroupName
         CmbGroup.ItemData(CmbGroup.NewIndex) = Asc(Left(!GroupID, 1)) & Asc(Mid(!GroupID, 2, 1)) & Asc(Mid(!GroupID, 3, 1))
         .MoveNext
      Loop
   End With
     
   CmbSubGroup.Clear
   With cn.Execute("Select * FROM SubGroups Order By SubGroupName")
      CmbSubGroup.AddItem "All SubGroups"
      CmbSubGroup.ItemData(CmbSubGroup.NewIndex) = 0
      Do Until .EOF
         CmbSubGroup.AddItem !SubGroupName
         CmbSubGroup.ItemData(CmbSubGroup.NewIndex) = !SubGroupID
         .MoveNext
      Loop
   End With
     
   CmbBrand.ListIndex = 0
   CmbCompany.ListIndex = 0
   CmbGroup.ListIndex = 0
   CmbSubGroup.ListIndex = 0
   CmbSortBy.ListIndex = 1
   'PopulateGrid

   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name <> Grid.Name Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnSelect.Enabled Then BtnSelect_Click
            KeyCode = 0
         Case vbKeyH
               FraHelp.ZOrder 0
               FraHelp.Visible = True
               KeyCode = 0
         Case vbKeyW
            If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtFromProductID.Name: If FunSelectFromProduct(ssFunctionKey, True) = True Then TxtToProductID.SetFocus
         Case TxtToProductID.Name: If FunSelectToProduct(ssFunctionKey, True) = True Then TxtProductName.SetFocus
      End Select
   End If
End Sub

Private Sub Grid_LostFocus()
   On Error GoTo ErrorHandler
   If vSuppressUpdateEvent Then Exit Sub
   If Grid.Visible = False Then Exit Sub
   UpdateRs
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  'If Grid.Columns(ColIndex).Text = "" Then Grid.Columns(ColIndex).Text = "0"
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   On Error GoTo ErrorHandler
   If vSuppressUpdateEvent Then Exit Sub
   UpdateRs
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_GotFocus()
   Grid.Row = 0
   Grid.Col = 0
   SendKeys "{Right}"
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event vbKeyRight, 1, 1, 1
      KeyCode = 0
   End If
End Sub

Private Sub UpdateRs()
   RsTemp.Filter = "ProductID = " & Val(Grid.Columns("ID").Text)
   If RsTemp.RecordCount = 0 And (Val(Grid.Columns("QtyLoose").Value) > 0 And Val(Grid.Columns("Price").Value) > 0) Then
      RsTemp.AddNew
      RsTemp!Productid = Val(Grid.Columns("ID").Text)
      RsTemp!ProductName = Grid.Columns("Name").Text
      RsTemp!QtyLoose = Val(Grid.Columns("QtyLoose").Text)
      RsTemp!Price = Val(Grid.Columns("Price").Text)
   ElseIf RsTemp.RecordCount = 1 And Val(Grid.Columns("QtyLoose").Value) = 0 And Val(Grid.Columns("Price").Value) = 0 Then
      RsTemp.Delete
   ElseIf RsTemp.RecordCount = 1 Then
      RsTemp!QtyLoose = Val(Grid.Columns("QtyLoose").Value)
      RsTemp!Price = Val(Grid.Columns("Price").Value)
      RsTemp.Update
  End If
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtFromProductID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If Trim(TxtFromProductID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectFromProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectFromProduct(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectFromProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectFromProduct = False: Exit Function
      TxtFromProductID.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtFromProductID.Text) = "" Then Exit Function
    If TxtFromProductID.Text = "" Then FunSelectFromProduct = False: Exit Function
    
   vStrSQL = " SELECT p.ProductID, Code, Qty, ProductName" & vbCrLf _
         + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
         + " where (p.productid = " & Val(TxtFromProductID.Text) & " or Code = '" & TxtFromProductID.Text & "')" & " and isLocked = 0 "

  With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtFromProductID.Text = !Productid
         FunSelectFromProduct = True
         .Close
         Exit Function
      Else
         FunSelectFromProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtFromProductID.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtToProductID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtToProductID.Text) = "" Then Exit Sub
   vTemp = Not FunSelectToProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectToProduct(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectToProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectToProduct = False: Exit Function
      TxtToProductID.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtToProductID.Text) = "" Then Exit Function
    If TxtToProductID.Text = "" Then FunSelectToProduct = False: Exit Function
    vStrSQL = " SELECT p.Productid, ProductName" & vbCrLf _
           + " from Products p" & vbCrLf _
           + " where p.productid = " & Val(TxtToProductID.Text)
  
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtToProductID.Text = !Productid
         FunSelectToProduct = True
         .Close
         Exit Function
      Else
         FunSelectToProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtToProductID.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

