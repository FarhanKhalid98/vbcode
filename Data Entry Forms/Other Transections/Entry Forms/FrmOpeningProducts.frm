VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmOpeningProducts 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmOpeningProducts.frx":0000
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
      Left            =   10830
      MaxLength       =   50
      TabIndex        =   22
      Top             =   3683
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.ComboBox CmbStatus 
      Height          =   315
      ItemData        =   "FrmOpeningProducts.frx":0ECA
      Left            =   7013
      List            =   "FrmOpeningProducts.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2513
      Width           =   1740
   End
   Begin VB.ComboBox CmbUsers 
      Height          =   315
      ItemData        =   "FrmOpeningProducts.frx":0ECE
      Left            =   5273
      List            =   "FrmOpeningProducts.frx":0ED0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2513
      Width           =   1740
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7770
      TabIndex        =   11
      Top             =   8483
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
      MICON           =   "FrmOpeningProducts.frx":0ED2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6450
      TabIndex        =   8
      Top             =   8483
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
      MICON           =   "FrmOpeningProducts.frx":0EEE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3810
      TabIndex        =   10
      Top             =   8483
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
      MICON           =   "FrmOpeningProducts.frx":0F0A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9090
      TabIndex        =   12
      Top             =   8483
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
      MICON           =   "FrmOpeningProducts.frx":0F26
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5130
      TabIndex        =   9
      Top             =   8483
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
      MICON           =   "FrmOpeningProducts.frx":0F42
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtID 
      Height          =   315
      Left            =   2925
      TabIndex        =   0
      Top             =   2513
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
      Left            =   3975
      TabIndex        =   1
      Top             =   2513
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
   Begin SITextBox.Txt TxtTotal 
      Height          =   315
      Left            =   8685
      TabIndex        =   7
      Top             =   7508
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   7
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   3735
      Left            =   4140
      TabIndex        =   6
      Top             =   3773
      Width           =   5745
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   3
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
      stylesets(0).Picture=   "FrmOpeningProducts.frx":0F5E
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
      stylesets(1).Picture=   "FrmOpeningProducts.frx":0F7A
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
      Columns.Count   =   3
      Columns(0).Width=   1852
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ProductID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   5900
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "ProductName"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1296
      Columns(2).Caption=   "Opening"
      Columns(2).Name =   "Opening"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   10134
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
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   8820
      TabIndex        =   4
      Tag             =   "NC"
      Top             =   2513
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   9495
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2513
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
      MICON           =   "FrmOpeningProducts.frx":0F96
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   9855
      TabIndex        =   23
      Tag             =   "NC"
      Top             =   2513
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
   Begin VB.Label Label1 
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
      Left            =   10845
      TabIndex        =   21
      Top             =   3413
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
      Left            =   9855
      TabIndex        =   20
      Top             =   2243
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
      Left            =   8820
      TabIndex        =   19
      Top             =   2243
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Index           =   2
      Left            =   7020
      TabIndex        =   18
      Top             =   2228
      Width           =   660
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Users"
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
      Index           =   1
      Left            =   5280
      TabIndex        =   17
      Top             =   2228
      Width           =   630
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   8055
      TabIndex        =   16
      Top             =   7553
      Width           =   555
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
      Left            =   2925
      TabIndex        =   15
      Top             =   2228
      Width           =   240
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
      Left            =   3975
      TabIndex        =   14
      Top             =   2228
      Width           =   1095
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Products"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   2700
      TabIndex        =   13
      Top             =   270
      Width           =   2535
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11610
      Top             =   60
      Width           =   375
   End
End
Attribute VB_Name = "FrmOpeningProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Public RsBody As New ADODB.Recordset
Dim vMode As FormMode
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.
Dim vid  As String
Dim sSql As String, vStrSQL As String, vCounter As Integer

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   If RsBody.State = adStateOpen Then
      RsBody.CancelBatch
      RsBody.Close
   End If
   Me.MousePointer = vbHourglass
   RsBody.Open "Select * From OpeningProductBody where ID = " & Val(TxtID.Text), cn, adOpenStatic, adLockBatchOptimistic
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   TxtTotal.Text = "0"
   sSql = " Select d.ProductID, d.ProductName, isnull(Opening,0) Opening " & vbCrLf _
      + " FROM (select * from OpeningProductHeader where ID = " & Val(TxtID.Text) & ")h " & vbCrLf _
      + " inner join OpeningProductBody b on h.ID = b.ID " & vbCrLf _
      + " right Outer Join (select * from Products where IsClosingProduct =1)d on b.ProductID = d.ProductID " & vbCrLf _
      + " Order by ProductName"
     
   With cn.Execute(sSql)
      Do Until .EOF
         Grid.AddNew
         Grid.Columns("ProductID").Value = .Fields("ProductID").Value
         Grid.Columns("ProductName").Value = .Fields("ProductName").Value
         Grid.Columns("Opening").Value = .Fields("Opening").Value
         TxtTotal.Text = Val(TxtTotal.Text) + Val(Grid.Columns("Opening").Value)
         Grid.Update
         .MoveNext
      Loop
   End With
   Grid.Redraw = True
   Grid.MoveFirst
   Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
   Call SubClearFields
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
   cn.Execute "Delete from OpeningProductBody where ID = " & Val(TxtID.Text)
   cn.Execute "Delete from OpeningProductHeader where ID = " & Val(TxtID.Text)
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   On Error GoTo ErrorHandler
   SchOpeningProduct.Show vbModal, Me
   If SchOpeningProduct.ParaOutID <> 0 Then GetOpeningProduct
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetOpeningProduct()
   On Error GoTo ErrorHandler
   'sSQL = "select * from OpeningProduct"
   'If Rs.State = adStateOpen Then Rs.Close
   'Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
   sSql = "Select H.ID, H.EntryDate, u.UserName, h.isVerify, h.StoreID, h.Tag, S.StoreName " & _
   " from OpeningProductHeader h inner join users u on h.ToUserNo = u.Userno " & _
   " left outer Join Stores s on S.StoreID = H.StoreID Where ID = " & SchOpeningProduct.ParaOutID
   With cn.Execute(sSql)
      If Not .BOF Then
          TxtID.Text = !ID
          DtpEntryDate.DateValue = !EntryDate
          TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
          TxtStoreName.Text = IIf(IsNull(!StoreName), "", !StoreName)
          TxtTag.Text = IIf(IsNull(!Tag), "", !Tag)
          CmbUsers.Text = !UserName
          CmbStatus.ListIndex = Abs(!isVerify)
      End If
      .Close
   End With
   FormStatus = OpenMode
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnStore_Click()
 If FunSelectStore(ssButton, False) = True Then
      TxtTotal.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub CmbStatus_Click()
   If CmbStatus.Visible = False Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub CmbUsers_Click()
   If CmbUsers.Visible = False Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "Opening Product"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   TxtStoreID.Text = ObjRegistry.StoreID
   FunSelectStore ssValidate, True
   TxtStoreID.Visible = ObjRegistry.StoreVisible
   BtnStore.Visible = ObjRegistry.StoreVisible
   TxtStoreName.Visible = ObjRegistry.StoreVisible
   LblStoreID.Visible = ObjRegistry.StoreVisible
   LblStoreName.Visible = ObjRegistry.StoreVisible
   With cn.Execute("Select * FROM Users where userno<>1")
      Do Until .EOF
         CmbUsers.AddItem !UserName
         CmbUsers.ItemData(CmbUsers.NewIndex) = !UserNo
         .MoveNext
      Loop
   End With
   CmbStatus.AddItem "Pending"
   CmbStatus.ItemData(CmbStatus.NewIndex) = 0
   CmbStatus.AddItem "Verified"
   CmbStatus.ItemData(CmbStatus.NewIndex) = 1
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
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then If TxtTotal.Enabled Then TxtTotal.SetFocus Else TxtStoreID.SetFocus
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
   If FunValidation = False Then Exit Sub
   cn.BeginTrans
   Set Rs = New ADODB.Recordset
   sSql = "Select * FROM OpeningProductHeader where ID = " & Val(TxtID.Text)
   Rs.Open sSql, cn, adOpenStatic, adLockOptimistic
   If vIsNewRecord Then
      Rs.AddNew
      Rs!ID = TxtID.Text
   End If
   Rs!EntryDate = DtpEntryDate.DateValue
   Rs!ToUserNo = CmbUsers.ItemData(CmbUsers.ListIndex)
   Rs!isVerify = CmbStatus.ItemData(CmbStatus.ListIndex)
   Rs!StoreID = TxtStoreID.Text
   Rs!Tag = IIf(Trim(TxtTag.Text) = "", "", TxtTag.Text)
   Rs!UserNo = vUser
   Rs.Update
   With RsBody
      .Filter = 0
      .MoveFirst
      For vCounter = 1 To .RecordCount
         !ID = Val(TxtID.Text)
         .MoveNext
      Next vCounter
      .UpdateBatch
   End With
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As String
   On Error GoTo ErrorHandler
   FunGetMaxID = cn.Execute("Select isnull(max(ID),0) + 1 from OpeningProductHeader").Fields(0)
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
   If TxtTotal.Enabled Then TxtTotal.SetFocus
   If Val(TxtTotal.Text) = "0" Then
       MsgBox "Please specify the One Opening", vbExclamation, "Alert"
       If Grid.Enabled And Grid.Visible Then Grid.SetFocus
       Exit Function
   End If
   If CmbUsers.ListIndex < 0 Then
       MsgBox "Please select user.", vbExclamation, "Alert"
       If CmbUsers.Enabled And CmbUsers.Visible Then CmbUsers.SetFocus
       Exit Function
   End If

'   If vIsNewRecord = True Then
'      Rs.Filter = " EntryDate = '" & DtpEntryDate.DateValue & "' and ToUserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)
'      If Rs.RecordCount <> 0 Then
'          MsgBox "This User Has ALready Petty Cash. Please specify the other User.", vbExclamation, "Alert"
'          CmbUsers.SetFocus
'          Exit Function
'      End If
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
      If CmbStatus.ListCount > 0 Then CmbStatus.ListIndex = 0
      LoadGrid
      vIsNewRecord = True
      If DtpEntryDate.Visible And DtpEntryDate.Enabled Then DtpEntryDate.SetFocus
   Case Is = OpenMode
      'TxtID.Enabled = False
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      DtpEntryDate.SetFocus
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
   Set FrmOpeningProducts = Nothing
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
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
   If ColIndex = 2 Then
      TxtTotal.Text = Val(TxtTotal.Text) + Val(Grid.Columns("Opening").Value) - Val(OldValue)
   End If
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   If Grid.Visible = False Then Exit Sub
   'If ActiveControl.Name <> Grid.Name Then Exit Sub
   'If Val(Grid.Columns("Qty").Value) = 0 Then Grid.Columns("Qty").Value = 0
   RsBody.Filter = "ID = " & Val(TxtID.Text) & " and ProductID = " & Val(Grid.Columns("ProductID").Value)
   If RsBody.RecordCount = 0 And Val(Grid.Columns("Opening").Value) > 0 Then
      RsBody.AddNew
      RsBody!ID = Val(TxtID.Text)
      RsBody!ProductID = Val(Grid.Columns("ProductID").Value)
      RsBody!Opening = Val(Grid.Columns("Opening").Value)
   ElseIf RsBody.RecordCount = 1 And Val(Grid.Columns("Opening").Value) = 0 Then
      RsBody.Delete
   ElseIf RsBody.RecordCount = 1 Then
      RsBody!Opening = Val(Grid.Columns("Opening").Value)
      RsBody.Update
   End If
End Sub

Private Sub Grid_Change()
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub Grid_GotFocus()
   Grid.Row = 0
   Grid.Col = 0
   SendKeys "{Right}"
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
          If BtnSave.Enabled = False And BtnSave.Visible = True Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectStore = False
          .Close
          TxtStoreID.Text = ""
          TxtStoreName.Text = ""
          If BtnSave.Enabled = False And BtnSave.Visible = True Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
