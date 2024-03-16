VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmApplyFixDiscounts 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmApplyFixDiscounts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtDiscPerC 
      Height          =   345
      Left            =   4230
      TabIndex        =   24
      Top             =   2205
      Width           =   795
   End
   Begin VB.TextBox TxtDiscPack 
      Height          =   345
      Left            =   13410
      TabIndex        =   22
      Top             =   2205
      Width           =   840
   End
   Begin VB.ComboBox CmbBrand 
      Height          =   315
      Left            =   12150
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1440
      Width           =   1890
   End
   Begin VB.ComboBox CmbSubGroup 
      Height          =   315
      Left            =   10410
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   1755
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   6855
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   1710
   End
   Begin VB.TextBox TxtDiscPer 
      Height          =   345
      Left            =   12600
      TabIndex        =   15
      Top             =   2205
      Width           =   795
   End
   Begin VB.TextBox TxtProductID 
      Height          =   345
      Left            =   6210
      TabIndex        =   12
      Top             =   2205
      Width           =   795
   End
   Begin JeweledBut.JeweledButton BtnFilter 
      Height          =   315
      Left            =   9990
      TabIndex        =   11
      Top             =   2205
      Width           =   900
      _ExtentX        =   1588
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
      MICON           =   "FrmApplyFixDiscounts.frx":0ECA
      BC              =   12632256
      FC              =   0
   End
   Begin VB.TextBox TxtProductName 
      Height          =   345
      Left            =   7020
      TabIndex        =   9
      Top             =   2220
      Width           =   2925
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   8595
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   1755
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5850
      Left            =   5899
      TabIndex        =   3
      Top             =   2565
      Width           =   8655
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   7
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
      stylesets(0).Picture=   "FrmApplyFixDiscounts.frx":0EE6
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
      stylesets(1).Picture=   "FrmApplyFixDiscounts.frx":0F02
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
      Columns.Count   =   7
      Columns(0).Width=   1402
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   5159
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2117
      Columns(2).Caption=   "Packing"
      Columns(2).Name =   "Packing"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).NumberFormat=   "########.##"
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   953
      Columns(3).Caption=   "Mul"
      Columns(3).Name =   "Mul"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   5
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   1693
      Columns(4).Caption=   "List Price"
      Columns(4).Name =   "ListPrice"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   5
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   1323
      Columns(5).Caption=   "Disc %"
      Columns(5).Name =   "DiscPer"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   5
      Columns(5).FieldLen=   256
      Columns(6).Width=   1561
      Columns(6).Caption=   "Disc Pack"
      Columns(6).Name =   "DiscPack"
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   5
      Columns(6).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   15266
      _ExtentY        =   10319
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
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6233
      TabIndex        =   4
      Top             =   8963
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
      MICON           =   "FrmApplyFixDiscounts.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7538
      TabIndex        =   5
      Top             =   8963
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
      MICON           =   "FrmApplyFixDiscounts.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10163
      TabIndex        =   6
      Top             =   8963
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
      MICON           =   "FrmApplyFixDiscounts.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnApply 
      Height          =   315
      Left            =   14325
      TabIndex        =   14
      Top             =   2205
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      TX              =   "Apply"
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
      MICON           =   "FrmApplyFixDiscounts.frx":0F72
      BC              =   12632256
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridC 
      Height          =   5850
      Left            =   807
      TabIndex        =   21
      Top             =   2565
      Width           =   5100
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   4
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
      stylesets(0).Picture=   "FrmApplyFixDiscounts.frx":0F8E
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
      stylesets(1).Picture=   "FrmApplyFixDiscounts.frx":0FAA
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
      Columns.Count   =   4
      Columns(0).Width=   1402
      Columns(0).Caption=   "Company ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   4075
      Columns(1).Caption=   "Company Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1323
      Columns(2).Caption=   "Disc %"
      Columns(2).Name =   "DiscPer"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   5
      Columns(2).FieldLen=   256
      Columns(3).Width=   1164
      Columns(3).Caption=   "DiscPack"
      Columns(3).Name =   "DiscPack"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   8996
      _ExtentY        =   10319
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
   Begin JeweledBut.JeweledButton BtnApplyC 
      Height          =   315
      Left            =   5085
      TabIndex        =   26
      Top             =   2205
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      TX              =   "Apply"
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
      MICON           =   "FrmApplyFixDiscounts.frx":0FC6
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc %"
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
      Left            =   4230
      TabIndex        =   25
      Top             =   1980
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc Pack"
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
      Left            =   13410
      TabIndex        =   23
      Top             =   1980
      Width           =   885
   End
   Begin VB.Label LblDepartment 
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
      Left            =   12150
      TabIndex        =   20
      Top             =   1170
      Width           =   1050
   End
   Begin VB.Label Label5 
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
      Left            =   10410
      TabIndex        =   18
      Top             =   1215
      Width           =   1395
   End
   Begin VB.Label Label4 
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
      Left            =   6855
      TabIndex        =   17
      Top             =   1215
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc %"
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
      Left            =   12600
      TabIndex        =   16
      Top             =   1980
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6210
      TabIndex        =   13
      Top             =   1935
      Width           =   930
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
      Left            =   7335
      TabIndex        =   10
      Top             =   1935
      Width           =   1215
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apply Fix Discounts"
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
      TabIndex        =   8
      Top             =   270
      Width           =   2625
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
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
      Left            =   8595
      TabIndex        =   7
      Top             =   1215
      Width           =   1065
   End
End
Attribute VB_Name = "FrmApplyFixDiscounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim RsC As New ADODB.Recordset
Dim vSuppressUpdateEvent As Boolean
Public ParaInVendorID As String
Dim ssql As String, vSQL As String, i As Integer

Private Sub BtnApply_Click()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   Grid.Redraw = False
   For i = 0 To Grid.Rows - 1
      Grid.Columns("DiscPer").Value = Val(TxtDiscPer.Text)
      Grid.Columns("DiscPack").Value = Val(TxtDiscPack.Text)
      Grid.MoveNext
   Next i
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnApplyC_Click()
   On Error GoTo ErrorHandler
   GridC.MoveFirst
   GridC.Redraw = False
   For i = 0 To Grid.Rows - 1
      GridC.Columns("DiscPer").Value = Val(TxtDiscPerC.Text)
      GridC.MoveNext
   Next i
   GridC.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbBrand_Click()
   On Error GoTo ErrorHandler
   If CmbGroup.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbGroup.Name Then Exit Sub
   Call SubFilter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbCompany_Click()
   On Error GoTo ErrorHandler
   If CmbCompany.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
   Call SubFilter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbGroup_Click()
   On Error GoTo ErrorHandler
   If CmbGroup.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbGroup.Name Then Exit Sub
   Call SubFilter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
   Call CmbGroup_Click
   'Call BtnFilter_Click
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   ' Product Update
   Grid.Update
   Rs.Filter = ""
   Rs.UpdateBatch
   ' Company Update
   GridC.Update
   RsC.Filter = ""
   RsC.UpdateBatch

   MsgBox "Your Entries has been Successfully Updated.", vbOKOnly + vbInformation, "Information"
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub SubFilter()
   ssql = " SELECT p.ProductID, ProductName, PackingName, Multiplier, ListPrice, isnull(v.DiscPer,0) as DiscPer, isnull(v.DiscPack,0) as DiscPack" & vbCrLf _
      + " FROM Products p left outer join packings pk on p.purchasepackingID = pk.packingID" & vbCrLf _
      + " left outer join ProductPacking pp on p.productid = pp.productid and p.purchasepackingID = pp.packingID" & vbCrLf _
      + " left outer join (select * from VendorProductDisc where VendorID = '" & ParaInVendorID & "' )v on p.productid = v.productid " & vbCrLf _
      + " where 1=1 " & IIf(CmbCompany.ListIndex = 0, "", " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex)) & IIf(CmbGroup.ListIndex = 0, "", " and GroupID ='" & GetGroupID(CmbGroup) & "'") & IIf(CmbSubGroup.ListIndex = 0, "", " and SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex)) & " Order by ProductName"
   Call PopulateGrid
End Sub

Private Sub BtnFilter_Click()
   ssql = " SELECT p.ProductID, ProductName, PackingName, Multiplier, ListPrice, isnull(v.DiscPer,0) as DiscPer, isnull(v.DiscPack,0) as DiscPack" & vbCrLf _
      + " FROM Products p left outer join packings pk on p.purchasepackingID = pk.packingID" & vbCrLf _
      + " left outer join ProductPacking pp on p.productid = pp.productid and p.purchasepackingID = pp.packingID" & vbCrLf _
      + " left outer join  (select * from VendorProductDisc where VendorID = '" & ParaInVendorID & "' )v on p.productid = v.productid " & vbCrLf _
      + " where 1=1 " & IIf(TxtProductID.Text = "", "", " and p.ProductID = '" & Right("00000" + CStr(Val(TxtProductID.Text)), 5) & "'") & IIf(Trim(TxtProductName.Text) = "", "", " and ProductName like '%" & TxtProductName.Text & "%'") & " Order by p.ProductID --ProductName"
      Call PopulateGrid
End Sub

Private Sub PopulateGrid()
   On Error GoTo ErrorHandler
   If Rs.State = adStateOpen Then
      Rs.CancelBatch
      Rs.Close
   End If
   Me.MousePointer = vbHourglass
   vSQL = "Select * FROM VendorProductDisc where VendorID = '" & ParaInVendorID & "'"
   Rs.Open vSQL, cn, adOpenStatic, adLockBatchOptimistic
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
   
   
   With cn.Execute(ssql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("Packing").Value = IIf(IsNull(!PackingName), "", !PackingName)
        Grid.Columns("Mul").Value = IIf(IsNull(!multiplier), "", !multiplier)
        Grid.Columns("ListPrice").Value = !ListPrice
        Grid.Columns("DiscPer").Value = !DiscPer
        Grid.Columns("DiscPack").Value = !DiscPack
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

Private Sub PopulateGridCC()
   On Error GoTo ErrorHandler
   If RsC.State = adStateOpen Then
      RsC.CancelBatch
      RsC.Close
   End If
   Me.MousePointer = vbHourglass
   vSQL = "Select * FROM VendorCompanyDisc where VendorID = '" & ParaInVendorID & "'"
   RsC.Open vSQL, cn, adOpenStatic, adLockBatchOptimistic
   GridC.Redraw = False
   GridC.CancelUpdate
   GridC.RemoveAll
   vSuppressUpdateEvent = True
   
   vSQL = " SELECT c.CompanyID, CompanyName, isnull(v.DiscPer,0) as DiscPer, isnull(v.DiscPack,0) as DiscPack" & vbCrLf _
      + " FROM Companies c " & vbCrLf _
      + " left outer join  (select * from VendorCompanyDisc where VendorID = '" & ParaInVendorID & "' )v on c.CompanyID = v.CompanyID " & vbCrLf _
      + " where 1=1 Order by CompanyName "

   
   With cn.Execute(vSQL)
      Do Until .EOF
        GridC.AddNew
        GridC.Columns("ID").Text = !companyid
        GridC.Columns("Name").Text = !CompanyName
        GridC.Columns("DiscPer").Value = !DiscPer
        GridC.Columns("DiscPack").Value = !DiscPack
        GridC.Update
        .MoveNext
      Loop
   End With
   vSuppressUpdateEvent = False
   GridC.Redraw = True
   GridC.MoveFirst
   Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   GridC.Redraw = True
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Function GetGroupID(cmb As ComboBox) As String
    On Error GoTo ErrorHandler
    If cmb.ListIndex <= 0 Then Exit Function
    GetGroupID = Chr(Left(cmb.ItemData(cmb.ListIndex), 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 3, 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 5, 2))
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub CmbSubGroup_Click()
   On Error GoTo ErrorHandler
   If CmbSubGroup.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbSubGroup.Name Then Exit Sub
   Call SubFilter
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Apply Fix Discounts"
   
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
   
   CmbCompany.ListIndex = 1
   CmbGroup.ListIndex = 0
   CmbSubGroup.ListIndex = 0
   CmbBrand.ListIndex = 0
   SubFilter
   
   PopulateGridCC
   
   'Call BtnFilter_Click
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name <> Grid.Name Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
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
      End Select
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then BtnSave.Enabled = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   On Error GoTo ErrorHandler
   If Val(Grid.Columns("DiscPer").Value) = 0 Then Grid.Columns("DiscPer").Value = 0
   If Val(Grid.Columns("DiscPack").Value) = 0 Then Grid.Columns("DiscPack").Value = 0
   Rs.Filter = " ProductID = '" & Grid.Columns("ID").Text & "'"
   If Rs.RecordCount = 0 And (Val(Grid.Columns("DiscPer").Value) > 0 Or Val(Grid.Columns("DiscPack").Value) > 0) Then
      Rs.AddNew
      Rs!VendorID = ParaInVendorID
      Rs!Productid = Grid.Columns("ID").Text
      Rs!DiscPer = Val(Grid.Columns("DiscPer").Value)
      Rs!DiscPack = Val(Grid.Columns("DiscPack").Value)
'      Rs!isChanged = 0
   ElseIf Rs.RecordCount = 1 And Val(Grid.Columns("DiscPer").Value) = 0 And Val(Grid.Columns("DiscPack").Value) = 0 Then
      Rs.Delete
   ElseIf Rs.RecordCount = 1 Then
      Rs!DiscPer = Val(Grid.Columns("DiscPer").Value)
      Rs!DiscPack = Val(Grid.Columns("DiscPack").Value)
'      Rs!isChanged = 1
      Rs.Update
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_Change()
   If BtnSave.Enabled = False Then BtnSave.Enabled = True
End Sub

Private Sub Grid_GotFocus()
   Grid.Row = 0
   Grid.Col = 0
'   SendKeys "{Right}"
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event vbKeyRight, 1, 1, 1
      KeyCode = 0
   End If
End Sub

Private Sub GridC_BeforeUpdate(Cancel As Integer)
   On Error GoTo ErrorHandler
   If Val(GridC.Columns("DiscPer").Value) = 0 Then GridC.Columns("DiscPer").Value = 0
   If Val(GridC.Columns("DiscPack").Value) = 0 Then GridC.Columns("DiscPack").Value = 0
   RsC.Filter = " CompanyID = '" & GridC.Columns("ID").Text & "'"
   If RsC.RecordCount = 0 And (Val(GridC.Columns("DiscPer").Value) > 0 Or Val(GridC.Columns("DiscPack").Value) > 0) Then
      RsC.AddNew
      RsC!VendorID = ParaInVendorID
      RsC!companyid = GridC.Columns("ID").Text
      RsC!DiscPer = Val(GridC.Columns("DiscPer").Value)
      RsC!DiscPack = Val(GridC.Columns("DiscPack").Value)
'      RsC!isChanged = 0
   ElseIf RsC.RecordCount = 1 And Val(GridC.Columns("DiscPer").Value) = 0 And Val(GridC.Columns("DiscPack").Value) = 0 Then
      RsC.Delete
   ElseIf RsC.RecordCount = 1 Then
      RsC!DiscPer = Val(GridC.Columns("DiscPer").Value)
      RsC!DiscPack = Val(GridC.Columns("DiscPack").Value)
'      RsC!isChanged = 1
      RsC.Update
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GridC_Change()
   If BtnSave.Enabled = False Then BtnSave.Enabled = True
End Sub

Private Sub GridC_GotFocus()
   GridC.Row = 0
   GridC.Col = 0
'   SendKeys "{Right}"
End Sub


Private Sub GridC_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event vbKeyRight, 1, 1, 1
      KeyCode = 0
   End If
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub
