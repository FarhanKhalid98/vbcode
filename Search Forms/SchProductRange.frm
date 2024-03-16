VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form SchProductRange 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnFromProduct 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4433
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   6038
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
      MICON           =   "SchProductRange.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtFromProductName 
      Height          =   315
      Left            =   4793
      TabIndex        =   6
      Tag             =   "nc"
      Top             =   6038
      Width           =   2985
      _ExtentX        =   5265
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
   Begin JeweledBut.JeweledButton BtnToProduct 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4448
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   6848
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
      MICON           =   "SchProductRange.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtToProductID 
      Height          =   315
      Left            =   3728
      TabIndex        =   1
      Top             =   6848
      Width           =   720
      _ExtentX        =   1270
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
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtToProductName 
      Height          =   315
      Left            =   4808
      TabIndex        =   8
      Tag             =   "nc"
      Top             =   6848
      Width           =   2985
      _ExtentX        =   5265
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
   Begin SITextBox.Txt TxtDefaultQty 
      Height          =   315
      Left            =   5251
      TabIndex        =   2
      Top             =   7793
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   8
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
      IntegralPoint   =   5
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   5768
      TabIndex        =   4
      Top             =   8565
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
      MICON           =   "SchProductRange.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   4463
      TabIndex        =   3
      Top             =   8565
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Select"
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
      MICON           =   "SchProductRange.frx":0054
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtFromProductID 
      Height          =   315
      Left            =   3713
      TabIndex        =   0
      Top             =   6038
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   15
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
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Range"
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
      Left            =   3000
      TabIndex        =   14
      Top             =   270
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Default Qty"
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
      Left            =   5266
      TabIndex        =   13
      Top             =   7553
      Width           =   990
   End
   Begin VB.Label Label7 
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
      Height          =   195
      Left            =   5183
      TabIndex        =   12
      Top             =   6623
      Width           =   1215
   End
   Begin VB.Label Label6 
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
      Height          =   195
      Left            =   3728
      TabIndex        =   11
      Top             =   6608
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Height          =   195
      Left            =   5363
      TabIndex        =   10
      Top             =   5798
      Width           =   1215
   End
   Begin VB.Label Label17 
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
      Height          =   195
      Left            =   3713
      TabIndex        =   9
      Top             =   5798
      Width           =   1395
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11378
      Top             =   2925
      Width           =   330
   End
End
Attribute VB_Name = "SchProductRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ParaOutFromID As String
Public ParaOutToID As String
Public ParaOutQty As Integer

Private Sub BtnFromProduct_Click()
   If FunSelectFromProduct(ssButton, True) = True Then
      TxtToProductID.SetFocus
   Else
      TxtFromProductID.SetFocus
   End If
End Sub

Private Sub BtnSelect_Click()
   If Trim(TxtFromProductID.Text) = "" Then
      MsgBox "Select From Product ID.", vbOKOnly + vbInformation, "Information"
   End If
   If Trim(TxtToProductID.Text) = "" Then
      MsgBox "Select To Product ID.", vbOKOnly + vbInformation, "Information"
   End If
   Me.ParaOutFromID = TxtFromProductID.Text
   Me.ParaOutToID = TxtToProductID.Text
   Me.ParaOutQty = TxtDefaultQty.Text
   Unload Me
End Sub

Private Sub BtnToProduct_Click()
   If FunSelectToProduct(ssButton, True) = True Then
      TxtDefaultQty.SetFocus
   Else
      TxtToProductID.SetFocus
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
          Case vbKeyQ
              If BtnClose.Enabled Then BtnClose_Click
              KeyCode = 0
          Case vbKeyS
              If BtnSelect.Enabled Then BtnSelect_Click
              KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtFromProductID.Name: If FunSelectFromProduct(ssFunctionKey, True) = True Then TxtToProductID.SetFocus
         Case TxtToProductID.Name: If FunSelectToProduct(ssFunctionKey, True) = True Then TxtDefaultQty.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   Me.ParaOutFromID = ""
   Me.ParaOutToID = ""
   Me.ParaOutQty = 0
   Unload Me
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Product Range"
   Me.MousePointer = vbDefault
   TxtDefaultQty.Text = "1"
   Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   Dim frmObj As Object
   For Each frmObj In Forms
       Set frmObj = Nothing
   Next
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtFromProductID_Change()
   If TxtFromProductID.Enabled = False Then Exit Sub
   If Me.ActiveControl.Name <> TxtFromProductID.Name Then Exit Sub
   If TxtFromProductName.Text <> "" Then
      TxtFromProductID.Text = ""
      TxtFromProductName.Text = ""
   End If
End Sub

Private Sub TxtFromProductID_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler
    If Me.ActiveControl.Name <> TxtFromProductID.Name Then Exit Sub
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
         TxtFromProductName.Text = !ProductName
         FunSelectFromProduct = True
         .Close
         Exit Function
      Else
         FunSelectFromProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtFromProductID.Text = ""
         TxtFromProductName.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtToProductID_Change()
   If TxtToProductID.Enabled = False Then Exit Sub
   If Me.ActiveControl.Name <> TxtToProductID.Name Then Exit Sub
   If TxtToProductName.Text <> "" Then TxtToProductName.Text = ""
End Sub

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
         TxtToProductName.Text = !ProductName
         FunSelectToProduct = True
         .Close
         Exit Function
      Else
         FunSelectToProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtToProductID.Text = ""
         TxtToProductName.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

