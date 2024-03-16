VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmDefineStockLimitStoreWise 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmDefineStockLimitStoreWise.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CmbSubGroup 
      Height          =   315
      Left            =   6773
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3113
      Width           =   1755
   End
   Begin VB.CheckBox ChkMinStockLimit 
      BackColor       =   &H80000009&
      Caption         =   "Min Stock Limit"
      Height          =   315
      Left            =   8528
      TabIndex        =   8
      Top             =   3113
      Width           =   1410
   End
   Begin VB.CheckBox ChkMaxStockLimit 
      BackColor       =   &H80000009&
      Caption         =   "Max Stock Limit"
      Height          =   315
      Left            =   9923
      TabIndex        =   9
      Top             =   3113
      Width           =   1455
   End
   Begin VB.ComboBox CmbCompany 
      Height          =   315
      Left            =   3308
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3113
      Width           =   1710
   End
   Begin VB.ComboBox CmbGroup 
      Height          =   315
      Left            =   5018
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3113
      Width           =   1755
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5370
      Left            =   3323
      TabIndex        =   14
      Top             =   3428
      Width           =   8715
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   5
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
      stylesets(0).Picture=   "FrmDefineStockLimitStoreWise.frx":0ECA
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
      stylesets(1).Picture=   "FrmDefineStockLimitStoreWise.frx":0EE6
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
      Columns.Count   =   5
      Columns(0).Width=   1852
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   5345
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1852
      Columns(2).Caption=   "Sale Qty"
      Columns(2).Name =   "SaleQty"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   5
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   2540
      Columns(3).Caption=   "Min Stock Limit"
      Columns(3).Name =   "MinStockLimit"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   5
      Columns(3).FieldLen=   256
      Columns(4).Width=   2566
      Columns(4).Caption=   "Max Stock Limit"
      Columns(4).Name =   "MaxStockLimit"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   5
      Columns(4).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   15372
      _ExtentY        =   9472
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
      Left            =   5573
      TabIndex        =   13
      Top             =   8993
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
      MICON           =   "FrmDefineStockLimitStoreWise.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   6893
      TabIndex        =   15
      Top             =   8993
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
      MICON           =   "FrmDefineStockLimitStoreWise.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8213
      TabIndex        =   16
      Top             =   8993
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
      MICON           =   "FrmDefineStockLimitStoreWise.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnApply 
      Height          =   315
      Left            =   11393
      TabIndex        =   12
      Top             =   3113
      Width           =   660
      _ExtentX        =   1164
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
      MICON           =   "FrmDefineStockLimitStoreWise.frx":0F56
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtMinStockLimit 
      Height          =   315
      Left            =   8483
      TabIndex        =   10
      Top             =   2798
      Width           =   1545
      _ExtentX        =   2725
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
   Begin SITextBox.Txt TxtMaxStockLimit 
      Height          =   315
      Left            =   10028
      TabIndex        =   11
      Top             =   2798
      Width           =   1380
      _ExtentX        =   2434
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
   Begin JeweledBut.JeweledButton BtnFilter 
      Height          =   315
      Left            =   10493
      TabIndex        =   23
      Top             =   1913
      Width           =   780
      _ExtentX        =   1376
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
      MICON           =   "FrmDefineStockLimitStoreWise.frx":0F72
      BC              =   12632256
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpTo 
      Height          =   315
      Left            =   9188
      TabIndex        =   4
      Top             =   1913
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   7883
      TabIndex        =   3
      Top             =   1913
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
      Left            =   3338
      TabIndex        =   0
      Tag             =   "NC"
      Top             =   1913
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Left            =   4478
      TabIndex        =   2
      Tag             =   "NC"
      Top             =   1913
      Width           =   2985
      _ExtentX        =   5265
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
      Left            =   4118
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1913
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
      MICON           =   "FrmDefineStockLimitStoreWise.frx":0F8E
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label19 
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
      Left            =   3323
      TabIndex        =   27
      Top             =   1718
      Width           =   720
   End
   Begin VB.Label Label18 
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
      Left            =   4478
      TabIndex        =   26
      Top             =   1718
      Width           =   1005
   End
   Begin VB.Label Label2 
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
      Left            =   9188
      TabIndex        =   25
      Top             =   1718
      Width           =   705
   End
   Begin VB.Label Label3 
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
      Left            =   7883
      TabIndex        =   24
      Top             =   1718
      Width           =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max Stock Limit"
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
      Left            =   10028
      TabIndex        =   22
      Top             =   2558
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min Stock Limit"
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
      Left            =   8483
      TabIndex        =   21
      Top             =   2558
      Width           =   1320
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
      Left            =   6773
      TabIndex        =   20
      Top             =   2888
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
      Left            =   3308
      TabIndex        =   19
      Top             =   2888
      Width           =   1320
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Define Stock Limit Store Wise"
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
      TabIndex        =   18
      Top             =   270
      Width           =   3900
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
      Left            =   5048
      TabIndex        =   17
      Top             =   2888
      Width           =   1065
   End
End
Attribute VB_Name = "FrmDefineStockLimitStoreWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vSuppressUpdateEvent As Boolean
Dim ssql As String, i As Integer

Private Sub BtnApply_Click()
   On Error GoTo ErrorHandler
   Grid.MoveFirst
   Grid.Redraw = False
   
   If ChkMaxStockLimit.Value = 1 And ChkMinStockLimit.Value = 0 Then
    For i = 0 To Grid.Rows - 1
      Grid.Columns("MaxStockLimit").Value = Val(Grid.Columns("SaleQty").Value)
      Grid.MoveNext
    Next i
   ElseIf ChkMinStockLimit.Value = 1 And ChkMaxStockLimit.Value = 0 Then
    For i = 0 To Grid.Rows - 1
      Grid.Columns("MinStockLimit").Value = Val(Grid.Columns("SaleQty").Value)
      Grid.MoveNext
    Next i
   ElseIf ChkMinStockLimit.Value = 1 And ChkMaxStockLimit.Value = 1 Then
    For i = 0 To Grid.Rows - 1
      Grid.Columns("MinStockLimit").Value = Val(Grid.Columns("SaleQty").Value)
      Grid.Columns("MaxStockLimit").Value = Val(Grid.Columns("SaleQty").Value)
      Grid.MoveNext
    Next i
    ElseIf ChkMinStockLimit.Value = 0 And ChkMaxStockLimit.Value = 0 Then
        If Trim(TxtMinStockLimit.Text) <> "" And Trim(TxtMaxStockLimit.Text) = "" Then
            For i = 0 To Grid.Rows - 1
                Grid.Columns("MinStockLimit").Value = Val(TxtMinStockLimit.Text)
                Grid.MoveNext
            Next i
        ElseIf Trim(TxtMaxStockLimit.Text) <> "" And Trim(TxtMinStockLimit.Text) = "" Then
            For i = 0 To Grid.Rows - 1
                Grid.Columns("MaxStockLimit").Value = Val(TxtMaxStockLimit.Text)
                Grid.MoveNext
            Next i
        ElseIf Trim(TxtMaxStockLimit.Text) <> "" And Trim(TxtMinStockLimit.Text) <> "" Then
            For i = 0 To Grid.Rows - 1
                Grid.Columns("MaxStockLimit").Value = Val(TxtMaxStockLimit.Text)
                Grid.Columns("MinStockLimit").Value = Val(TxtMinStockLimit.Text)
                Grid.MoveNext
            Next i
        End If
   End If
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnFilter_Click()
   On Error GoTo ErrorHandler
   'If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
Abc:
  If Rs.State = adStateOpen Then
    Rs.CancelBatch
    Rs.Close
  End If
  Me.MousePointer = vbHourglass
  ssql = "Select * FROM StockLimitStoreWise"
  Rs.Open ssql, cn, adOpenStatic, adLockBatchOptimistic
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
    ssql = "Select p.ProductID, ProductName, isnull(prodSale.SaleQty,0) as SaleQty, RetailPrice, isnull(SL.MaxStockLimit,0) MaxStockLimit, isnull(SL.MinSTockLimit,0) as MinSTockLimit, GroupID, CompanyID, SubgroupID From " & vbCrLf _
        + " ( " & vbCrLf _
        + " select ProductID,  Sum(SaleQty) - sum(ReturnQty) as SaleQTy " & vbCrLf _
        + " From " & vbCrLf _
        + " ( " & vbCrLf _
        + " Select ProductID, Sum(isnull(qty,0)) as SaleQty, 0 as ReturnQty" & vbCrLf _
        + " from SaleBody b inner join SaleHeader h on h.billID = b.billID and h.billDate = b.BillDate" & vbCrLf _
        + " Where h.billdate between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "' and h.StoreID = " & Val(TxtStoreID.Text) & vbCrLf _
        + " Group By ProductID" & vbCrLf _
        + " Union All" & vbCrLf _
        + " Select ProductID, 0, Sum(isnull(qty,0)) as ReturnQty" & vbCrLf _
        + " from SaleReturnBody b inner join SaleReturnHeader h on h.ReturnID = b.ReturnID and h.ReturnDate = b.ReturnDate" & vbCrLf _
        + " Where h.Returndate between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "' and h.StoreID = " & Val(TxtStoreID.Text) & vbCrLf _
        + " Group By ProductID" & vbCrLf _
        + " )d Group by ProductID" & vbCrLf _
        + " )prodSale " & vbCrLf _
        + " Right Outer Join Products P on P.productID = ProdSale.ProductID" & vbCrLf _
        + " Left Outer Join StockLimitStoreWise SL on Sl.ProductID = P.ProductID " & vbCrLf _
        + " where 1=1 " & IIf(CmbCompany.ListIndex = 0, "", " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex)) & IIf(CmbGroup.ListIndex = 0, "", " and GroupID ='" & GetGroupID(CmbGroup) & "'") & IIf(CmbSubGroup.ListIndex = 0, "", " and SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex)) & " Order by ProductName"
   With cn.Execute(ssql)
      Do Until .EOF
         Grid.AddNew
         Grid.Columns("ID").Text = !Productid
         Grid.Columns("Name").Text = !ProductName
         Grid.Columns("SaleQty").Value = !SaleQty
         Grid.Columns("MaxStockLimit").Value = !MaxStockLimit
         Grid.Columns("MinStockLimit").Value = !MinStockLimit
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
   If Err.Number = 91 Then GoTo Abc
   Grid.Redraw = True
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnStore_Click()
 On Error GoTo ErrorHandler
   If FunSelectStore(ssButton, False) = True Then
      DtpFrom.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbCompany_Click()
   On Error GoTo ErrorHandler
   If CmbCompany.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbCompany.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbGroup_Click()
   On Error GoTo ErrorHandler
   If CmbGroup.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbGroup.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
   Call BtnFilter_Click
   'Call BtnFilter_Click
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   Grid.Update
'   Rs.MoveFirst
'   While Not Rs.EOF
'      If Rs.EditMode <> adEditNone Then
'         Call ActivityLog("Change Price", eEdit, , , Rs!ProductID)
'      End If
'      Rs.MoveNext
'   Wend
   Rs.Filter = ""
   Rs.UpdateBatch
   MsgBox "Your Entries has been Successfully Updated.", vbOKOnly + vbInformation, "Information"
   Exit Sub
ErrorHandler:
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

Private Sub PopulateGrid()
   On Error GoTo ErrorHandler
     If Rs.State = adStateOpen Then
     Rs.CancelBatch
     Rs.Close
   End If
   Me.MousePointer = vbHourglass
   ssql = "Select * FROM StockLimitStoreWise"
   Rs.Open ssql, cn, adOpenStatic, adLockBatchOptimistic
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
  ssql = " Select p.ProductID, ProductName, isnull(prodSale.SaleQty,0) as SaleQty, RetailPrice, isnull(SL.MaxStockLimit,0) MaxStockLimit, isnull(SL.MinSTockLimit,0) as MinSTockLimit, GroupID, CompanyID, SubgroupID From " & vbCrLf _
        + " ( " & vbCrLf _
        + " select ProductID,  Sum(SaleQty) - sum(ReturnQty) as SaleQty " & vbCrLf _
        + " From " & vbCrLf _
        + " ( " & vbCrLf _
        + " Select ProductID, Sum(isnull(qty,0)) as SaleQty, 0 as ReturnQty" & vbCrLf _
        + " from SaleBody b inner join SaleHeader h on h.billID = b.billID and h.billDate = b.BillDate" & vbCrLf _
        + " Where h.billdate between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "' and h.StoreID = " & Val(TxtStoreID.Text) & vbCrLf _
        + " Group By ProductID" & vbCrLf _
        + " Union All" & vbCrLf _
        + " Select ProductID, 0, Sum(isnull(qty,0)) as ReturnQty" & vbCrLf _
        + " from SaleReturnBody b inner join SaleReturnHeader h on h.ReturnID = b.ReturnID and h.ReturnDate = b.ReturnDate" & vbCrLf _
        + " Where h.Returndate between '" & DtpFrom.DateValue & "' and '" & DtpTo.DateValue & "' and h.StoreID = " & Val(TxtStoreID.Text) & vbCrLf _
        + " Group By ProductID" & vbCrLf _
        + " )d Group by ProductID " & vbCrLf _
        + " )prodSale " & vbCrLf _
        + " Right Outer Join Products P on P.productID = ProdSale.ProductID" & vbCrLf _
        + " Left Outer Join StockLimitStoreWise SL on Sl.ProductID = P.ProductID " & vbCrLf _
        + " where 1=1 " & IIf(CmbCompany.ListIndex = 0, "", " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex)) & IIf(CmbGroup.ListIndex = 0, "", " and GroupID ='" & GetGroupID(CmbGroup) & "'") & IIf(CmbSubGroup.ListIndex = 0, "", " and SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex)) & " Order by ProductName"
'  sSql = "  SELECT p.ProductID, ProductName, RetailPrice, " & vbCrLf _
'        + " isnull(SL.MinStockLimit,0) as MinStockLimit , isnull(SL.MaxStockLimit,0) as MaxStockLimit " & vbCrLf _
'        + " FROM Products p Left Outer join StockLimitStoreWise SL on SL.ProductID = p.ProductID " & vbCrLf _
'        + " where 1=1  " & IIf(CmbCompany.ListIndex = 0, "", " and CompanyID =" & CmbCompany.ItemData(CmbCompany.ListIndex)) & IIf(CmbGroup.ListIndex = 0, "", " and GroupID ='" & GetGroupID(CmbGroup) & "'") & IIf(CmbSubGroup.ListIndex = 0, "", " and SubGroupID =" & CmbSubGroup.ItemData(CmbSubGroup.ListIndex)) & " Order by ProductName"
   With cn.Execute(ssql)
      Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !Productid
        Grid.Columns("Name").Text = !ProductName
        Grid.Columns("SaleQty").Value = !SaleQty
        Grid.Columns("MinStockLimit").Value = !MinStockLimit
        Grid.Columns("MaxStockLimit").Value = !MaxStockLimit
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

Private Sub CmbSubGroup_Click()
   On Error GoTo ErrorHandler
   If CmbSubGroup.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbSubGroup.Name Then Exit Sub
   Call PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Define Stock Limit Store Wise"
   
   DtpFrom.DateValue = Date
   DtpTo.DateValue = Date
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
      
   CmbCompany.ListIndex = 1
   CmbGroup.ListIndex = 0
   CmbSubGroup.ListIndex = 0
   PopulateGrid
'  Grid.Columns("Name").Locked = Not ObjUserSecurity.IsAdministrator
'  If CmbCompany.ListCount > 0 Then CmbCompany.ListIndex = 0
'  CmbGroup.ListIndex = 0
   
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
      ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         'Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, True) = True Then TxtProductID.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  'If Grid.Columns(ColIndex).Text = "" Then Grid.Columns(ColIndex).Text = "0"
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   On Error GoTo ErrorHandler
   'If Grid.Visible = False Then Exit Sub
   'If ActiveControl.Name <> Grid.Name Then Exit Sub
   If Val(Grid.Columns("MinStockLimit").Value) = 0 Then Grid.Columns("MinStockLimit").Value = 0
   If Val(Grid.Columns("MaxStockLimit").Value) = 0 Then Grid.Columns("MaxStockLimit").Value = 0
   Rs.Filter = " ProductID = " & Val(Grid.Columns("ID").Text)
   If Rs.RecordCount = 0 And Val(Grid.Columns("MinStockLimit").Value) > 0 Then
      Rs.AddNew
      Rs!Productid = Grid.Columns("ID").Text
      Rs!StoreID = Val(TxtStoreID.Text)
      Rs!MinStockLimit = Val(Grid.Columns("MinStockLimit").Value)
      Rs!MaxStockLimit = Val(Grid.Columns("MaxStockLimit").Value)
   ElseIf Rs.RecordCount = 1 Then
      Rs!MinStockLimit = Val(Grid.Columns("MinStockLimit").Value)
      Rs!MaxStockLimit = Val(Grid.Columns("MaxStockLimit").Value)
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

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo ErrorHandler
      Set FrmDefineStockLimitStoreWise = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtStoreID_Change()
   If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then
      TxtStoreName.Text = ""
   End If
End Sub

Private Sub TxtStoreID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   If Me.ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
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
