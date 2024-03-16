VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmUserClosing 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmUserClosing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   10406
      MaxLength       =   50
      TabIndex        =   7
      Top             =   3218
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox TxtTotalWithoutPettyCash 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10271
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   7448
      Width           =   1995
   End
   Begin VB.TextBox TxtTotalPettyCash 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5366
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   7493
      Width           =   1995
   End
   Begin VB.TextBox TxtTotalCash 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7571
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   7493
      Width           =   1995
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   3735
      Left            =   3296
      TabIndex        =   0
      Top             =   3278
      Width           =   6465
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
      stylesets(0).Picture=   "FrmUserClosing.frx":0ECA
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
      stylesets(1).Picture=   "FrmUserClosing.frx":0EE6
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedCol"
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
      SelectTypeRow   =   0
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   106
      Columns.Count   =   8
      Columns(0).Width=   2117
      Columns(0).Caption=   "Denomination"
      Columns(0).Name =   "Denom"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   714
      Columns(1).Name =   "Mul"
      Columns(1).Alignment=   2
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2778
      Columns(2).Caption=   "Petty Cash Quantity"
      Columns(2).Name =   "PQty"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1296
      Columns(3).Caption=   "Quantity"
      Columns(3).Name =   "Qty"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   794
      Columns(4).Name =   "Equ"
      Columns(4).Alignment=   2
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   2646
      Columns(5).Caption=   "Amount"
      Columns(5).Name =   "Amount"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "PAmount"
      Columns(6).Name =   "PAmount"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "QAmount"
      Columns(7).Name =   "QAmount"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   11404
      _ExtentY        =   6588
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
      Left            =   3296
      TabIndex        =   2
      Top             =   2318
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpEntryDate 
      Height          =   315
      Left            =   4346
      TabIndex        =   3
      Top             =   2318
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
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7054
      TabIndex        =   12
      Top             =   8633
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Remove"
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
      MICON           =   "FrmUserClosing.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   5734
      TabIndex        =   13
      Top             =   8633
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save"
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
      MICON           =   "FrmUserClosing.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3094
      TabIndex        =   14
      Top             =   8633
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Open"
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
      MICON           =   "FrmUserClosing.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8374
      TabIndex        =   15
      Top             =   8633
      Width           =   1275
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
      MICON           =   "FrmUserClosing.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4414
      TabIndex        =   16
      Top             =   8633
      Width           =   1275
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
      MICON           =   "FrmUserClosing.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   5681
      TabIndex        =   4
      Tag             =   "NC"
      Top             =   2318
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
      Left            =   6716
      TabIndex        =   6
      Tag             =   "NC"
      Top             =   2318
      Width           =   1395
      _ExtentX        =   2461
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
      Left            =   6356
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2318
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
      MICON           =   "FrmUserClosing.frx":0F8E
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tag"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10406
      TabIndex        =   23
      Top             =   2978
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6716
      TabIndex        =   22
      Top             =   2078
      Width           =   1245
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5681
      TabIndex        =   21
      Top             =   2078
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Without Petty Cash"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9720
      TabIndex        =   20
      Top             =   7185
      Width           =   2550
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cash"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8246
      TabIndex        =   17
      Top             =   7178
      Width           =   1305
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Petty Cash"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5366
      TabIndex        =   10
      Top             =   7178
      Width           =   1995
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4346
      TabIndex        =   9
      Top             =   2078
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3296
      TabIndex        =   8
      Top             =   2078
      Width           =   240
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Closing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   2700
      TabIndex        =   1
      Top             =   270
      Width           =   1785
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "FrmUserClosing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim RsBody As New ADODB.Recordset
Dim vSuppressUpdateEvent As Boolean
Dim vPosition As Byte
Dim vTaskKey As String
Dim vRowCounter As Integer, vCounter As Integer
Dim vMode As FormMode
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.
Dim vid As String
Dim sSql As String, vStrSQL As String
Dim vExcess, vShort As Double

'Private Sub LoadGrid1()
'   On Error GoTo ErrorHandler
'   If RsBody.State = adStateOpen Then
'      RsBody.CancelBatch
'      RsBody.Close
'   End If
'   Me.MousePointer = vbHourglass
'   RsBody.Open "Select * From UserClosingBody where ID = " & Val(TxtID.Text), CN, adOpenStatic, adLockBatchOptimistic
'   Grid.Redraw = False
'   Grid.CancelUpdate
'   Grid.RemoveAll
'   vSuppressUpdateEvent = True
'   TxtTotalCash.Text = "0"
'   With CN.Execute("Select d.Denom, isnull(Qty,0) Qty  FROM (select * from UserClosingHeader where id = " & Val(TxtID.Text) & ")h inner join UserClosingBody b on h.ID = b.ID right Outer Join (select * from Denominations)d on b.Denom = d.Denom Order by d.Denom desc")
'      Do Until .EOF
'         Grid.AddNew
'         Grid.Columns("Denom").Value = .Fields("Denom").Value
'         Grid.Columns("Mul").Text = "X"
'         Grid.Columns("Equ").Text = "="
'         Grid.Columns("Qty").Value = .Fields("Qty").Value
'         Grid.Columns("Amount").Value = Val(.Fields("Denom").Value) * (.Fields("Qty").Value)
'         TxtTotalCash.Text = Val(TxtTotalCash.Text) + Val(.Fields("Denom").Value) * Val(.Fields("Qty").Value)
'         Grid.Update
'         .MoveNext
'      Loop
'   End With
'   vSuppressUpdateEvent = False
'   Grid.Redraw = True
'   Grid.MoveFirst
'   'If Grid.Visible Then Grid.SetFocus
'   Me.MousePointer = vbDefault
'   Exit Sub
'ErrorHandler:
'   Grid.Redraw = True
'   Me.MousePointer = vbDefault
'   Call ShowErrorMessage
'End Sub

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   If RsBody.State = adStateOpen Then
      RsBody.CancelBatch
      RsBody.Close
   End If
   Me.MousePointer = vbHourglass
   RsBody.Open "Select * From UserClosingBody where ID = " & Val(TxtID.Text), cn, adOpenStatic, adLockBatchOptimistic
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
   TxtTotalCash.Text = "0"
   TxtTotalPettyCash.Text = "0"
   'sSQL = "Select d.Denom, isnull(Qty,0) Qty  FROM (select * from UserClosingHeader where UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex) & " and EntryDate = '" & DtpEntryDate.DateValue & "')h inner join UserClosingBody b on h.ID = b.ID right Outer Join Denominations d on b.Denom = d.Denom Order by d.Denom desc"
   If vIsNewRecord = True Then
      sSql = " Select d.Denom, isnull(h.totalcash,0) totalcash, isnull(PQty,0) PQty, isnull(Qty,0) Qty  " & vbCrLf _
            + " FROM (select * from UserClosingHeader where EntryDate = '" & DtpEntryDate.DateValue & "' and UserNo = " & vUser & " )h " & vbCrLf _
            + " inner join UserClosingBody b on h.ID = b.ID " & vbCrLf _
            + " right Outer Join Denominations d on b.Denom = d.Denom " & vbCrLf _
            + " Order by d.Denom desc"
   Else
      sSql = " Select d.Denom, isnull(h.totalcash,0) totalcash, isnull(PQty,0) PQty, isnull(Qty,0) Qty  " & vbCrLf _
         + " FROM (select * from UserClosingHeader where ID = " & Val(TxtID.Text) & ")h " & vbCrLf _
         + " inner join UserClosingBody b on h.ID = b.ID " & vbCrLf _
         + " right Outer Join Denominations d on b.Denom = d.Denom " & vbCrLf _
         + " Order by d.Denom desc"
     
   End If
   With cn.Execute(sSql)
      Do Until .EOF
         Grid.AddNew
         Grid.Columns("Denom").Value = .Fields("Denom").Value
         Grid.Columns("Mul").Text = "X"
         Grid.Columns("Equ").Text = "="
         Grid.Columns("PQty").Value = .Fields("PQty").Value
         Grid.Columns("Qty").Value = .Fields("Qty").Value
         Grid.Columns("PAmount").Value = Val(.Fields("Denom").Value) * (.Fields("PQty").Value)
         Grid.Columns("QAmount").Value = Val(.Fields("Denom").Value) * (.Fields("Qty").Value)
         Grid.Columns("Amount").Value = Val(Grid.Columns("PAmount").Value) + (Grid.Columns("QAmount").Value)
'         TxtTotalCash.Text = .Fields("TotalCash").Value
         TxtTotalCash.Text = Val(TxtTotalCash.Text) + Val(Grid.Columns("Amount").Value)
         TxtTotalPettyCash.Text = Val(TxtTotalPettyCash.Text) + Val(Grid.Columns("PAmount").Value)
         TxtTotalWithoutPettyCash.Text = Val(TxtTotalCash.Text) - Val(TxtTotalPettyCash.Text)
         Grid.Update
         .MoveNext
      Loop
   End With
'   TxtTotalWithoutPettyCash.Text = Val(TxtTotalCash.Text) - Val(TxtTotalPettyCash.Text)
   
   
   vSuppressUpdateEvent = False
   Grid.Redraw = True
   Grid.MoveFirst
   'If Grid.Visible Then Grid.SetFocus
   Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
   'Call SubClearFields
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
   Dim vtbl As String
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
   cn.BeginTrans
   cn.Execute "Delete from UserClosingBody where ID = " & Val(TxtID.Text)
   cn.Execute "Delete from UserClosingHeader where ID = " & Val(TxtID.Text)
   cn.Execute "Delete from AdminClosing where UserNo = " & vUser & " and Entrydate = '" & DtpEntryDate.DateValue & "'"
   Call SetNonPost(vUser, DtpEntryDate.DateValue)
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   On Error GoTo ErrorHandler
   SchUserClosing.Show vbModal, Me
   If SchUserClosing.ParaOutID <> 0 Then GetUserClosing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetUserClosing()
   On Error GoTo ErrorHandler
   'sSQL = "select * from AdminClosing"
   'If Rs.State = adStateOpen Then Rs.Close
   'Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
   sSql = " Select H.ID, H.EntryDate, h.StoreID, h.Tag, S.StoreName " & _
   " from UserClosingHeader H Left Outer Join Stores s on S.StoreID = H.StoreID where ID = " & SchUserClosing.ParaOutID
   With cn.Execute(sSql)
      If Not .BOF Then
          TxtID.Text = !ID
          DtpEntryDate.DateValue = !EntryDate
          TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
          TxtStoreName.Text = IIf(IsNull(!StoreName), "", !StoreName)
          TxtTag.Text = IIf(IsNull(!Tag), "", !Tag)
      End If
      .Close
   End With
   FormStatus = OpenMode
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubInitilize()
   On Error GoTo ErrorHandler
   With cn.Execute("Select * from UserClosingHeader where EntryDate = '" & DtpEntryDate.DateValue & "' and UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtID.Text = !ID
         DtpEntryDate.DateValue = !EntryDate
         FormStatus = OpenMode
      End If
   End With
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubCheck()
   On Error GoTo ErrorHandler
   If vIsNewRecord = True Then
      If cn.Execute("Select * from UserClosingHeader where EntryDate = '" & DtpEntryDate.DateValue & "' and UserNo = " & vUser).RecordCount > 0 Then
         MsgBox "You have Already Enter Closing of Current Date. Please specify other Date.", vbExclamation, "Alert"
         DtpEntryDate.SetFocus
         Exit Sub
      Else
         LoadGrid
      End If
   End If
   If vIsNewRecord = False And cn.Execute("Select * from UserClosingHeader where EntryDate = '" & DtpEntryDate.DateValue & "' and UserNo = " & vUser & " and isposted = 1").RecordCount > 0 Then
      MsgBox "You have Already Enter Closing of Current Date. Please specify other Date.", vbExclamation, "Alert"
      DtpEntryDate.SetFocus
      Exit Sub
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnStore_Click()
 If FunSelectStore(ssButton, False) = True Then
      TxtTotalPettyCash.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub DtpEntryDate_Change()
   On Error GoTo ErrorHandler
   Call SubCheck
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "User Closing"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   
   TxtStoreID.Text = ObjRegistry.StoreID
   
   With cn.Execute("select * from UserRegistry where UserNo = " & vUser)
      If .RecordCount > 0 Then
         TxtStoreID.Text = IIf(IsNull(!StoreID), ObjRegistry.StoreID, !StoreID)
      End If
      .Close
   End With
   
   FunSelectStore ssValidate, True
   
   
   TxtStoreID.Visible = ObjRegistry.StoreVisible
   BtnStore.Visible = ObjRegistry.StoreVisible
   TxtStoreName.Visible = ObjRegistry.StoreVisible
   LblStoreID.Visible = ObjRegistry.StoreVisible
   LblStoreName.Visible = ObjRegistry.StoreVisible
   
   FormStatus = NewMode
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name <> Grid.Name Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then If TxtTotalPettyCash.Enabled Then TxtTotalPettyCash.SetFocus Else TxtStoreID.SetFocus
      End Select
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
         Case vbKeyO
            If BtnOpen.Enabled Then BtnOpen_Click
            KeyCode = 0
         Case vbKeyR
            If BtnDelete.Enabled Then BtnDelete_Click
            KeyCode = 0
      End Select
   Else
      If UCase(Me.ActiveControl.Name) Like "TXT*" Or UCase(Me.ActiveControl.Name) Like "DTP*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
      Exit Sub
   End If
   
   If cn.Execute("select * from AdminClosing where ToUserNo = " & vUser & " and EntryDate = '" & DtpEntryDate.DateValue & "' and StoreID = " & TxtStoreID.Text).RecordCount > 0 Then
      MsgBox "This User on that Date have Already Closing. Please specify other.", vbExclamation, "Alert"
      Exit Sub
   End If
   
   Grid.Update
   RsBody.Filter = ""
   If FunValidation = False Then Exit Sub
   cn.BeginTrans
   
   'If vIsNewRecord = False Then Call ActivityLog("Sale Invoice", eEdit, TxtBillID.Text, DtpBillDate.DateValue)
   sSql = "Select * from UserClosingHeader where ID = " & Val(TxtID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open sSql, cn, adOpenStatic, adLockOptimistic
      If .BOF Then
         .AddNew
         !ID = Val(TxtID.Text)
      End If
      !EntryDate = DtpEntryDate.DateValue
      !TotalCash = TxtTotalCash.Text
      !StoreID = TxtStoreID.Text
      !Tag = IIf(Trim(TxtTag.Text) = "", "", TxtTag.Text)
      !UserNo = vUser
      !isPosted = 0
      .Update
      .Close
   End With
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !ID = Val(TxtID.Text)
         !StoreID = Val(TxtStoreID.Text)
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   
   If ObjRegistry.AdminClosingSaveWhenUserClosingSaved Then Call SaveAdminClosing
     
   cn.CommitTrans
   MsgBox "Record has been Saved Successfully.", vbOKOnly + vbInformation, "Information"
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As String
   On Error GoTo ErrorHandler
   FunGetMaxID = cn.Execute("Select isnull(max(ID),0) + 1 from UserClosingHeader").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
Private Function FunGetAdminClosingMaxID() As String
   On Error GoTo ErrorHandler
   FunGetAdminClosingMaxID = cn.Execute("Select isnull(max(ID),0) + 1 from AdminClosing where StoreID = " & Val(TxtStoreID.Text)).Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
Private Function FunValidation() As Boolean
   On Error GoTo ErrorHandler
   If Trim(TxtID.Text) = "" Then
      MsgBox "Please specify ID", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
   End If
   Call Grid_BeforeUpdate(False)
   'If TxtAmount.Enabled Then TxtAmount.SetFocus
   If vIsNewRecord = True Then
      If cn.Execute("Select * from UserClosingHeader where EntryDate = '" & DtpEntryDate.DateValue & "' and UserNo = " & vUser).RecordCount > 0 Then
         MsgBox "You have Already Enter Closing of Current Date. Please specify other Date.", vbExclamation, "Alert"
         DtpEntryDate.SetFocus
         Exit Function
      End If
      If cn.Execute("Select * from UserClosingHeader where ID = " & Val(TxtID.Text)).RecordCount > 0 Then
         TxtID.Text = FunGetMaxID
      End If
   End If
   If vIsNewRecord = False And cn.Execute("Select * from UserClosingHeader where EntryDate = '" & DtpEntryDate.DateValue & "' and UserNo = " & vUser & " and isposted = 1").RecordCount > 0 Then
      MsgBox "You have Already Enter Closing of Current Date. Please specify other Date.", vbExclamation, "Alert"
      DtpEntryDate.SetFocus
      Exit Function
   End If
   
   
'   If CN.Execute("Select * from UserClosingHeader where UserNo = " & ObjUserSecurity.UserNo & " and EntryDate = '" & DtpEntryDate.DateValue & "' and isposted = 1").RecordCount = 0 Then
'      With CN.Execute("Select * from UserClosingHeader where UserNo = " & ObjUserSecurity.UserNo & " and EntryDate = '" & DtpEntryDate.DateValue & "' and isposted = 0")
'         If .RecordCount > 0 Then
'            TxtID.Text = !ID
'         Else
'            TxtID.Text = CN.Execute("select isnull(max(ID),0)+1 from UserClosingHeader").Fields(0).Value
'         End If
'      End With
'   End If
   
   'All Ok, now validation is success
  
   FunValidation = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Property Get FormStatus() As FormMode
   On Error GoTo ErrorHandler
   'Nothing
   FormStatus = vMode
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
   'Based upon the value of vNewValue, we shall decide what controls to enable/disable
   On Error GoTo ErrorHandler
   vMode = vNewValue
   Select Case vNewValue
   Case Is = NewMode
      Call SubClearFields
      'TxtID.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      DtpEntryDate.DateValue = IIf(Format(Now, "hh") > 3, Date, DateAdd("d", -1, Date))
      TxtID.Text = FunGetMaxID
      vIsNewRecord = True
      Call SubInitilize
      If DtpEntryDate.Visible And DtpEntryDate.Enabled Then DtpEntryDate.SetFocus
   Case Is = OpenMode
      'TxtID.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      If DtpEntryDate.Visible And DtpEntryDate.Enabled Then DtpEntryDate.SetFocus
      vIsNewRecord = False
   Case Is = ChangeMode
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo ErrorHandler
   Set RptReportViewer.Report = Nothing
   Set FrmUserClosing = Nothing
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
      ElseIf TypeOf ctl Is TextBox Then
         If ctl.Tag = "" Then
            ctl.Text = ""
         End If
      End If
   Next
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
   If ColIndex = 3 Then
      TxtTotalCash.Text = Val(TxtTotalCash.Text) + Val(Grid.Columns("QAmount").Value) - (Val(OldValue) * Val(Grid.Columns("Denom").Value))
   ElseIf ColIndex = 2 Then
      TxtTotalPettyCash.Text = Val(TxtTotalPettyCash.Text) + Val(Grid.Columns("PAmount").Value) - (OldValue * Val(Grid.Columns("Denom").Value))
      TxtTotalCash.Text = Val(TxtTotalCash.Text) + Grid.Columns("PAmount").Value - (Val(OldValue) * Val(Grid.Columns("Denom").Value))
   End If
   TxtTotalWithoutPettyCash.Text = Val(TxtTotalCash.Text) - Val(TxtTotalPettyCash.Text)
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   If Grid.Visible = False Then Exit Sub
   'If ActiveControl.Name <> Grid.Name Then Exit Sub
   'If Val(Grid.Columns("Qty").Value) = 0 Then Grid.Columns("Qty").Value = 0
   RsBody.Filter = "ID=" & Val(TxtID.Text) & " and Denom=" & Grid.Columns("Denom").Value
   If RsBody.RecordCount = 0 And (Val(Grid.Columns("Qty").Value) > 0 Or Val(Grid.Columns("PQty").Value) > 0) Then
      RsBody.AddNew
      RsBody!ID = Val(TxtID.Text)
      RsBody!Denom = Val(Grid.Columns("Denom").Value)
      RsBody!Qty = Val(Grid.Columns("Qty").Value)
      RsBody!PQty = Val(Grid.Columns("PQty").Value)
   ElseIf RsBody.RecordCount = 1 And Val(Grid.Columns("Qty").Value) = 0 And Val(Grid.Columns("PQty").Value) = 0 Then
      RsBody.Delete
   ElseIf RsBody.RecordCount = 1 Then
      RsBody!Qty = Val(Grid.Columns("Qty").Value)
      RsBody!PQty = Val(Grid.Columns("PQty").Value)
      RsBody.Update
   End If
End Sub

Private Sub Grid_Change()
   Grid.Columns("PAmount").Value = Val(Grid.Columns("Denom").Value) * Val(Grid.Columns("PQty").Value)
   Grid.Columns("QAmount").Value = Val(Grid.Columns("Denom").Value) * Val(Grid.Columns("Qty").Value)
   Grid.Columns("Amount").Value = Val(Grid.Columns("PAmount").Value) + Grid.Columns("QAmount").Value
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub Grid_GotFocus()
   Grid.Row = 0
   Grid.Col = 0
'   SendKeys "{Right}"
End Sub

Private Sub Grid_LostFocus()
   Call Grid_BeforeUpdate(False)
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtStoreID_Change()
If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then TxtStoreName.Text = ""
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

Private Sub SaveAdminClosing()
On Error GoTo ErrorHandler

   Call CalculateAmount(vUser, DtpEntryDate.DateValue)
   
   If (Val(TxtTotalCash.Text) - Val(vCashAvailable)) > 0 Then
      vExcess = (Val(TxtTotalCash.Text) - Val(vCashAvailable))
      vShort = 0
   Else
      vShort = (Val(TxtTotalCash.Text) - Val(vCashAvailable))
      vExcess = 0
   End If
   
   sSql = "Insert into AdminClosing (ID,EntryDate,TotalSale,PettyCash,BankCardSale,CreditSale,Discount,SaleReturn,TotalCash,AddCollection,ToUserNo,UserNo,RecoveryCustomer,Payment,StoreID,CashReceived,ServiceCharges,STax,Excess,Short) " & vbCrLf & _
          " Values (" & FunGetAdminClosingMaxID & ",'" & DtpEntryDate.DateValue & "'," & vTotalSale & "," & vPettyCash & "," & vBankCardSale & "," & vCreditSale & "," & vDiscount & "," & vSaleReturn & "," & Val(TxtTotalCash.Text) & "," & 0 & "," & vUser & "," & vUser & "," & vRecoveryCustomer & "," & vPayments & "," & TxtStoreID.Text & "," & vCashReceived & "," & vServiceCharges & "," & vSTax & "," & IIf(vExcess = 0, "Null", vExcess) & "," & IIf(vShort = 0, "Null", vShort) & ")"
   
   cn.Execute (sSql)
   
   Call SetPost(vUser, DtpEntryDate.DateValue)
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
