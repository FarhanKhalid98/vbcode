VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form DefSProducts 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15450
   Icon            =   "DefSProducts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   770
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkLockProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Lock Product"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8295
      TabIndex        =   6
      Top             =   6113
      Width           =   1320
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
      Height          =   4290
      Left            =   11385
      TabIndex        =   25
      Top             =   1125
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
         Height          =   3855
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         Tag             =   "NC"
         Text            =   "DefSProducts.frx":0ECA
         Top             =   330
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
         TabIndex        =   27
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.TextBox TxtFilterID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3375
      TabIndex        =   11
      Top             =   2858
      Width           =   2655
   End
   Begin VB.TextBox TxtFilterProductName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3375
      TabIndex        =   10
      Top             =   2453
      Width           =   2655
   End
   Begin SITextBox.Txt TxtPurPrice 
      Height          =   315
      Left            =   8340
      TabIndex        =   2
      Top             =   3848
      Width           =   1080
      _ExtentX        =   1905
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
      DecimalPoint    =   2
      IntegralPoint   =   6
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4095
      Left            =   2505
      TabIndex        =   12
      Top             =   3353
      Width           =   4050
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
      stylesets(0).Picture=   "DefSProducts.frx":0FC8
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
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
      Columns.Count   =   3
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "GroupID"
      Columns(0).Name =   "GroupID"
      Columns(0).DataField=   "Column 2"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1667
      Columns(1).Caption=   "Product ID"
      Columns(1).Name =   "ID"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 0"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4392
      Columns(2).Caption=   "Name"
      Columns(2).Name =   "Name"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 1"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   7144
      _ExtentY        =   7223
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
      Left            =   2400
      TabIndex        =   13
      Top             =   8678
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
      MICON           =   "DefSProducts.frx":0FE4
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   3720
      TabIndex        =   14
      Top             =   8678
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
      MICON           =   "DefSProducts.frx":1000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   5040
      TabIndex        =   15
      Top             =   8678
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
      MICON           =   "DefSProducts.frx":101C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   8640
      TabIndex        =   7
      Top             =   8678
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
      MICON           =   "DefSProducts.frx":1038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9960
      TabIndex        =   8
      Top             =   8678
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
      MICON           =   "DefSProducts.frx":1054
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   11280
      TabIndex        =   9
      Top             =   8678
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
      MICON           =   "DefSProducts.frx":1070
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtRetailPrice 
      Height          =   315
      Left            =   8340
      TabIndex        =   3
      Top             =   4373
      Width           =   1080
      _ExtentX        =   1905
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
      DecimalPoint    =   2
      IntegralPoint   =   6
   End
   Begin SITextBox.Txt TxtID 
      Height          =   315
      Left            =   8340
      TabIndex        =   0
      Tag             =   "nc"
      Top             =   3248
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   5
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
      IntegralPoint   =   5
   End
   Begin SITextBox.Txt TxtName 
      Height          =   315
      Left            =   9465
      TabIndex        =   1
      Top             =   3248
      Width           =   3585
      _ExtentX        =   6324
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
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtPurDisc 
      Height          =   315
      Left            =   8340
      TabIndex        =   4
      Top             =   4883
      Width           =   1080
      _ExtentX        =   1905
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
      Masked          =   2
      DecimalPoint    =   3
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtSaleDisc 
      Height          =   315
      Left            =   8340
      TabIndex        =   5
      Top             =   5408
      Width           =   1080
      _ExtentX        =   1905
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
      Masked          =   2
      DecimalPoint    =   3
      IntegralPoint   =   3
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Disc/Pc"
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
      Left            =   7230
      TabIndex        =   28
      Top             =   4943
      Width           =   1050
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
      Left            =   11160
      TabIndex        =   24
      Top             =   585
      Width           =   435
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Disc/Pc"
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
      Left            =   7140
      TabIndex        =   23
      Top             =   5483
      Width           =   1140
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Products"
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
      TabIndex        =   22
      Top             =   270
      Width           =   2250
   End
   Begin VB.Label Label1 
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
      Height          =   195
      Left            =   7350
      TabIndex        =   21
      Top             =   3293
      Width           =   930
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code :"
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
      Left            =   2760
      TabIndex        =   20
      Top             =   2903
      Width           =   570
   End
   Begin VB.Image ImgExit 
      Height          =   300
      Left            =   11625
      Top             =   45
      Width           =   345
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   461
      X2              =   461
      Y1              =   168.533
      Y2              =   496.533
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
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
      Left            =   2715
      TabIndex        =   19
      Top             =   2498
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retail Price"
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
      Left            =   7260
      TabIndex        =   18
      Top             =   4433
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pur Price"
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
      Left            =   7425
      TabIndex        =   17
      Top             =   3893
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   8970
      TabIndex        =   16
      Top             =   3308
      Width           =   495
   End
End
Attribute VB_Name = "DefSProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim vMode As FormMode
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.
Dim RsCode As New ADODB.Recordset
Dim Flag As Boolean
Dim vPer As Byte
Dim vProductID As String
Dim vMaxBinID As Integer
Dim vCounter As Integer

Private Sub ChkLockProduct_Click()
   If ActiveControl.Name <> ChkLockProduct.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   Else
      Dim frmObj As Object
      For Each frmObj In Forms
          Set frmObj = Nothing
      Next
      Set Rs = Nothing
      Set Rs1 = Nothing
      Set DefSProducts = Nothing
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub FraHelp_Click()
   FraHelp.Visible = False
End Sub

Private Sub Grid_DblClick()
   If Grid.Rows > 0 And BtnOpen.Enabled Then BtnOpen_Click
End Sub

Private Sub CmbFilterGroup_Click()
  On Error GoTo ErrorHandler
    Set Rs1 = New ADODB.Recordset
    Rs1.Open "Select * FROM SProducts Order By ProductName", cn, adOpenStatic, adLockOptimistic
    Set Grid.DataSource = Rs1
    Grid.Columns("ID").DataField = "ProductID"
    Grid.Columns("Name").DataField = "ProductName"
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      KeyCode = 0
   ElseIf KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
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
         Case vbKeyN
             If BtnNew.Enabled Then BtnNew_Click
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
      End Select
   ElseIf KeyCode = vbKeyF9 Then
      vPer = Val(InputBox("Add Percentage in Purchase Price to Convert Retail Price", "Input", vPer))
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('Service Products'" & "," & Val(vProductID) & ",Null,'Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  FormStatus = SelectionMode
End Sub

Private Sub BtnClose_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('Service Products'" & "," & Val(vProductID) & ",Null,'Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Unload Me
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniServiceProducts", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
  Dim vtbl As String
  Dim vProductID As String
  If Rs1.RecordCount > 0 Then
    If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
    Dim vid As String
    vid = Rs1!Productid
    vtbl = Common.ChildDataExists("SProducts", "ProductId='" & vid & "'", "", "ProductID")
    If vtbl <> "" Then
      MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
      Exit Sub
    End If
    '---------------------------------------------
    Call ActivityLog("SProducts", eDelete, , , vid)
    
    vMaxBinID = FunGetMaxBinID
    ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
    cn.Execute ("Insert Into Bin_Sproducts Select " & vMaxBinID & ",'" & Date & "',* from Sproducts Where productID = " & TxtID.Text)

    vProductID = TxtID.Text  'TxtPrefix.Text & TxtID.Text
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('Service Products'" & "," & vProductID & ",Null,'Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Rs1.Delete
    Rs1.ReQuery
    Rs.ReQuery
    '---------------------------------------------
    If Rs1.RecordCount = 0 Then FormStatus = NewMode: Exit Sub
    Rs1.MoveNext
    Grid.MoveNext
    If Rs1.EOF Then Rs1.MoveLast
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnNew_Click()
  FormStatus = NewMode
End Sub

Private Sub BtnOpen_Click()
  On Error GoTo ErrorHandler
  If Rs1.RecordCount > 0 Then
    If Rs1.BOF = False And Rs1.EOF = False Then
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
   vUserAction = UserAuthentication("MniServiceProducts", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   Call UserActivities
   
   'Rs.Filter = ""
   Rs.Filter = "ProductID='" & TxtID.Text & "'"
   If vIsNewRecord = False Then Call ActivityLog("SProducts", eEdit, , , TxtID.Text)
   If vIsNewRecord And Rs.RecordCount = 0 Then
      Rs.AddNew
      Rs!Productid = TxtID.Text
   End If
   If ObjUserSecurity.IsAdministrator = True Then Rs!PurPrice = Val(TxtPurPrice.Text)
   Rs!ProductName = TxtName.Text
   Rs!RetailPrice = Val(TxtRetailPrice.Text)
   Rs!PurPrice = Val(TxtPurPrice.Text)
   Rs!DiscPC = IIf(Val(TxtSaleDisc.Text) = 0, 0, TxtSaleDisc.Text)
   Rs!PurDiscPC = IIf(Val(TxtPurDisc.Text) = 0, 0, TxtPurDisc.Text)
   Rs!IsLocked = ChkLockProduct.Value
   Rs.Update
   Rs.ReQuery
   Rs1.ReQuery
   If vIsNewRecord = True Then Call ActivityLog("SProducts", eAdd, , , TxtID.Text)
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
  If TxtID.Enabled = True And cn.Execute("select count(*) from SProducts where productid = '" & TxtID.Text & "'").Fields(0) > 0 Then
    MsgBox "This ID already exists. A new ID has been generated. Please save again", vbExclamation, "Alert"
    TxtID.Text = FunGetMaxID
    TxtID.SetFocus
    Exit Function
  End If
  If vIsNewRecord Then
    If Trim(TxtID.Text) = "" Then
      MsgBox "Please specify a Product ID", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
    If Not IsNumeric(TxtID.Text) Then
      MsgBox "The Product ID must be numeric", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
  End If
  If Trim(TxtName.Text) = "" Then
    MsgBox "Please specify a Product Name", vbExclamation, "Alert"
    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
    Exit Function
  End If
  
'    If Val(TxtPurchasePrice.Text) = 0 Then
'      MsgBox "The Purchase Price must be Greater than zero", vbExclamation, "Alert"
'      If TxtPurchasePrice.Enabled And TxtPurchasePrice.Visible Then TxtPurchasePrice.SetFocus
'      Exit Function
'    ElseIf Val(TxtSalePrice.Text) = 0 Then
'      MsgBox "The Sale price must be greater than zero", vbExclamation, "Alert"
'      If TxtSalePrice.Enabled And TxtSalePrice.Visible Then TxtSalePrice.SetFocus
'      Exit Function
'    ElseIf Val(TxtPurchaseDiscRatio.Text) > 99.99 Then
'      MsgBox "The Purchase Disc.(%) must be less than 99.99", vbExclamation, "Alert"
'      If TxtPurchaseDiscRatio.Enabled And TxtPurchaseDiscRatio.Visible Then TxtPurchaseDiscRatio.SetFocus
'      Exit Function
'    ElseIf Val(TxtPurchaseDiscRatio.Text) > 0 And Val(TxtPurchaseDiscVal.Text) > 0 Then
'      MsgBox "Only one of the Purchase Disc (%) or Purchase Disc. value must be provided.", vbExclamation, "Alert"
'      If TxtPurchaseDiscRatio.Enabled And TxtPurchaseDiscRatio.Visible Then TxtPurchaseDiscRatio.SetFocus
'      Exit Function
'    ElseIf Val(TxtSaleDiscRatio.Text) > 99.99 Then
'      MsgBox "The Sale Disc.(%) must be less than 99.99", vbExclamation, "Alert"
'      If TxtSaleDiscRatio.Enabled And TxtSaleDiscRatio.Visible Then TxtSaleDiscRatio.SetFocus
'      Exit Function
'    ElseIf Val(TxtSaleDiscRatio.Text) > 0 And Val(TxtSaleDiscVal.Text) > 0 Then
'      MsgBox "Only one of the Sale Disc (%) or Sale Disc. value must be provided.", vbExclamation, "Alert"
'      If TxtSaleDiscRatio.Enabled And TxtSaleDiscRatio.Visible Then TxtSaleDiscRatio.SetFocus
'      Exit Function
'    ElseIf Val(TxtPurchaseSTRatio.Text) > 99.99 Then
'      MsgBox "The Purchase S-Tax (%) must be less than 99.99", vbExclamation, "Alert"
'      If TxtPurchaseSTRatio.Enabled And TxtPurchaseSTRatio.Visible Then TxtPurchaseSTRatio.SetFocus
'      Exit Function
'    ElseIf Val(TxtSaleSTRatio.Text) > 99.99 Then
'      MsgBox "The Sale S-Tax (%) must be less than 99.99", vbExclamation, "Alert"
'      If TxtSaleSTRatio.Enabled And TxtSaleSTRatio.Visible Then TxtSaleSTRatio.SetFocus
'      Exit Function
'    End If
'
  'All Ok, now validation is success
  FunValidation = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub Form_Load()
  On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Products"
   HelpLocation Me
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "Select * FROM SProducts order by ProductName", cn, adOpenStatic, adLockOptimistic
    vPer = 0
    CmbFilterGroup_Click
    FormStatus = NewMode
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
      'If Val(TxtID.Text) <> 0 Then
      '   TxtID.Text = Right("00000" + CStr(Val(TxtID.Text) + 1), 5)
      'Else
         TxtID.Text = FunGetMaxID
      'End If
      Call SubClearFields(True)
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtFilterProductName.Enabled = False
      TxtFilterID.Enabled = False
      Grid.Enabled = False
      If TxtName.Visible And TxtName.Enabled Then TxtName.SetFocus
      vIsNewRecord = True
    Case Is = OpenMode
      Call SubClearFields(True)
      Call Grid_RowColChange(0, 0)
      'TxtGroupName.Enabled = False
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnClear.Enabled = True
      Grid.Enabled = False
      TxtID.Enabled = False
      vIsNewRecord = False
    Case Is = ChangeMode
      BtnSave.Enabled = True
    Case Is = SelectionMode
      'CmbFilterGroup.ListIndex = 0
      Call SubClearFields(False)
      Call Grid_RowColChange(0, 0)
      Grid.Enabled = True
      TxtFilterProductName.Enabled = True
      TxtFilterID.Enabled = True
      Grid.SetFocus
      BtnNew.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnSave.Enabled = False
      BtnClear.Enabled = False
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub Grid_Click()
    If Grid.Rows > 0 Then Call Grid_RowColChange(0, 0)
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKeyA To vbKeyZ, Asc("a") To Asc("z")
      TxtFilterProductName.Text = TxtFilterProductName.Text & Chr(KeyAscii): TxtFilterProductName.SelStart = Len(TxtFilterProductName.Text): TxtFilterProductName.SetFocus
   Case vbKey0 To vbKey9
      TxtFilterID.Text = TxtFilterID.Text & Chr(KeyAscii): TxtFilterID.SelStart = Len(TxtFilterID.Text): TxtFilterID.SetFocus
   End Select
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo ErrorHandler
   If Trim(Grid.Columns("ID").Text) <> "" Then
      If Rs1.RecordCount > 0 And Grid.Enabled Then
         TxtID.Text = Grid.Columns("ID").Text
         TxtName.Text = Grid.Columns("Name").Text
         '''''''''''''''''''''''''''''''''''''''''''''
         TxtPurPrice.Text = IIf(ObjUserSecurity.IsAdministrator = True, Rs1!PurPrice, 0)
         TxtRetailPrice.Text = Rs1!RetailPrice
         TxtPurDisc.Text = IIf(IsNull(Rs1!PurDiscPC), 0, Rs1!PurDiscPC)
         TxtSaleDisc.Text = IIf(IsNull(Rs1!DiscPC), 0, Rs1!DiscPC)
         ChkLockProduct.Value = Abs(Rs1!IsLocked)
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
      If TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then ctl.Text = ""
         If ctl.Tag = "" Then ctl.Enabled = Enable
      ElseIf TypeOf ctl Is TextBox Then
         If ctl.Tag = "" Then ctl.Text = ""
      ElseIf TypeOf ctl Is ComboBox Then
         If ctl.Tag = "" Then ctl.Enabled = Enable
      ElseIf TypeOf ctl Is JeweledButton Then
         If ctl.Tag <> "" Then ctl.Enabled = Enable
      ElseIf TypeOf ctl Is CheckBox Then
         ctl.Value = 0
         ctl.Enabled = Enable
      End If
   Next
   Rs.Filter = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As String
   On Error GoTo ErrorHandler
   FunGetMaxID = cn.Execute("Select right('00000' + cast(isnull(max(cast(ProductId as smallint)),0) + 1 as varchar),5) from SProducts ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub ImgExit_Click()
   Unload Me
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

Private Sub TxtFilterID_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtFilterID.Name Then Exit Sub
   If Trim(TxtFilterID.Text) = "" Then Grid.MoveFirst: Exit Sub
   If Len(TxtFilterID.Text) > 5 Then
      Rs1.Find "ProductID ='" & Right("00000" + CStr(Val(TxtFilterID.Text)), 5) & "'", , adSearchForward, 1
   End If
   If Rs1.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub TxtID_LostFocus()
   If Len(TxtID.Text) = 5 Then Exit Sub
   TxtID.Text = Right("00000" + CStr(Val(TxtID.Text)), 5)
End Sub

Private Sub TxtPurPrice_Change()
   If ActiveControl.Name <> TxtPurPrice.Name Then Exit Sub
   If vPer = 0 Then Exit Sub
   TxtRetailPrice.Text = SelfRound(Val(TxtPurPrice.Text) + (Val(TxtPurPrice.Text) * vPer / 100))
End Sub

Private Sub TxtFilterProductName_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtFilterProductName.Name Then Exit Sub
   If Trim(TxtFilterProductName.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs1.Find "ProductName like '" & Replace(TxtFilterProductName.Text, "'", "''") & "%'", , adSearchForward, 1
   If Rs1.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtName_Change()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtName.Name Then Exit Sub
   'If Trim(TxtName.Text) = "" Then Exit Sub
   Set Rs1 = New ADODB.Recordset
   Rs1.Open "Select * FROM SProducts where ProductName like '%" & Replace(TxtName.Text, "'", "''") & "%' Order By ProductName", cn, adOpenStatic, adLockOptimistic
   Set Grid.DataSource = Rs1
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxBinID = cn.Execute("Select isnull(max(BinID),0)+1 from Bin_Sproducts ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
    If vIsNewRecord = False Then
    With cn.Execute("Select  * from SProducts where ProductID =" & TxtID.Text)
        If TxtName.Text <> !ProductName Then
            cn.Execute ("Insert Into UserActivities values ('Service Products'" & "," & TxtID.Text & ", Null , 'Updated Product Name-" & !ProductName & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtPurPrice.Text <> IIf(IsNull(!PurPrice), "", !PurPrice) Then
            cn.Execute ("Insert Into UserActivities values ('Service Products'" & "," & TxtID.Text & ", Null , 'Updated PurPrice-" & !PurPrice & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtPurDisc.Text <> IIf(IsNull(!PurDiscPC), "", !PurDiscPC) Then
            cn.Execute ("Insert Into UserActivities values ('Service Products'" & "," & TxtID.Text & ", Null , 'Updated PurDiscPC-" & !PurDiscPC & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtSaleDisc.Text <> IIf(IsNull(!DiscPC), "", !DiscPC) Then
            cn.Execute ("Insert Into UserActivities values ('Service Products'" & "," & TxtID.Text & ", Null , 'Updated DiscPC-" & !DiscPC & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
    End With
   Else
        cn.Execute ("Insert Into UserActivities values ('Service Products'" & "," & TxtID.Text & ", Null ,'Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub
