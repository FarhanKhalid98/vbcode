VERSION 5.00
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmPriceChecker 
   BorderStyle     =   0  'None
   Caption         =   "Price Checker"
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   Icon            =   "FrmPriceChecker.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin SITextBox.Txt TxtDiscVal 
      Height          =   315
      Left            =   9945
      TabIndex        =   5
      Top             =   2265
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   6
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
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   1020
      TabIndex        =   0
      Top             =   2265
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      Appearance      =   0
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
      IntegralPoint   =   15
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   6690
      TabIndex        =   1
      Top             =   2265
      Width           =   780
      _ExtentX        =   1376
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
   Begin SITextBox.Txt TxtPrice 
      Height          =   315
      Left            =   7470
      TabIndex        =   2
      Top             =   2265
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      Enabled         =   0   'False
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
      Left            =   11625
      TabIndex        =   6
      Top             =   2265
      Width           =   1650
      _ExtentX        =   2910
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
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   3195
      TabIndex        =   12
      Top             =   2265
      Width           =   3495
      _ExtentX        =   6165
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
   Begin SITextBox.Txt TxtDiscPC 
      Height          =   315
      Left            =   8430
      TabIndex        =   3
      Top             =   2265
      Width           =   825
      _ExtentX        =   1455
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
   Begin SITextBox.Txt TxtDiscPer 
      Height          =   315
      Left            =   9255
      TabIndex        =   4
      Top             =   2265
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   6
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
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtActualAmount 
      Height          =   315
      Left            =   13275
      TabIndex        =   17
      Top             =   2265
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
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
   Begin SITextBox.Txt TxtSC 
      Height          =   315
      Left            =   10935
      TabIndex        =   19
      Top             =   2265
      Width           =   690
      _ExtentX        =   1217
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
   Begin VB.Label LblDiscountedPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Discounted Price"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   840
      Left            =   8190
      TabIndex        =   28
      Top             =   9360
      Width           =   5880
   End
   Begin VB.Label LblOriginalPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Original Price"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   840
      Left            =   2430
      TabIndex        =   27
      Top             =   9360
      Width           =   4680
   End
   Begin VB.Label TxtOriginalPrice 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4425
      Left            =   1290
      TabIndex        =   26
      Top             =   4770
      Width           =   6450
   End
   Begin VB.Label LblAllStock 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Store Stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   12645
      TabIndex        =   25
      Top             =   1530
      Width           =   1905
   End
   Begin VB.Label LblStock 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   13005
      TabIndex        =   24
      Top             =   1170
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label LblStockCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   13095
      TabIndex        =   23
      Top             =   810
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TxtProductName2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   """"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   840
      Left            =   7620
      TabIndex        =   22
      Top             =   3195
      Width           =   360
   End
   Begin VB.Label TxtNetAmount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4380
      Left            =   7860
      TabIndex        =   21
      Top             =   4770
      Width           =   6450
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "S.C."
      Height          =   195
      Left            =   10935
      TabIndex        =   20
      Top             =   2070
      Width           =   300
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Amount"
      Height          =   195
      Left            =   13305
      TabIndex        =   18
      Top             =   2070
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price Checker"
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
      Left            =   2700
      TabIndex        =   16
      Top             =   270
      Width           =   2370
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. %"
      Height          =   195
      Left            =   9255
      TabIndex        =   15
      Top             =   2070
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   195
      Left            =   7470
      TabIndex        =   14
      Top             =   2070
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3195
      TabIndex        =   13
      Top             =   2070
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   1020
      TabIndex        =   11
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc / PC"
      Height          =   195
      Left            =   8430
      TabIndex        =   10
      Top             =   2070
      Width           =   690
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   6690
      TabIndex        =   9
      Top             =   2070
      Width           =   240
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   11625
      TabIndex        =   8
      Top             =   2070
      Width           =   540
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. Val"
      Height          =   195
      Left            =   9945
      TabIndex        =   7
      Top             =   2070
      Width           =   630
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
      Begin VB.Menu MniCostPrice 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmPriceChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vQtyloose As Double
Dim vShowDiscPrice As Boolean
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error GoTo ErrorHandler
  If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case TxtCode.Name
         If FunSelectProduct(ssValidate, False) = True Then Exit Sub
      Case Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End Select
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   LblAllStock.Visible = False
   vShowDiscPrice = ObjRegistry.ShowDiscPrice
   If vShowDiscPrice Then
      TxtNetAmount.Visible = True
      TxtOriginalPrice.Visible = False
      LblOriginalPrice.Visible = False
      LblDiscountedPrice.Visible = False
      TxtNetAmount.Left = 201
      TxtNetAmount.Top = 384
      TxtNetAmount.Width = 626
   End If
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
End Sub

Private Sub TxtCode_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtCode.Text) = "" Then Exit Sub
   vTemp = Not FunSelectProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectProduct(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   
    '---------------------------
   If TxtCode.Enabled = False Then FunSelectProduct = False: Exit Function
   If Trim(TxtCode.Text) = "" Then FunSelectProduct = False: Exit Function
    
   If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
   
    vStrSQL = " SELECT p.productid, code, Qty, ProductName, ServiceCharges, RetailPrice, DiscPer, DiscPC, EmpComm, TokenVal, isChangedPrice" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where (p.productid = " & TxtCode.Text & " or code='" & TxtCode.Text & "')" & " and isLocked = 0 "

   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductName.Text = !ProductName
         vIsChangedPrice = !isChangedPrice
         TxtPrice.Text = !RetailPrice
         vUnitPrice = Val(TxtPrice.Text)
         TxtQty.Text = IIf(Len(TxtCode.Text) <= 5 And IsNumeric(TxtCode.Text), 1, IIf(IsNull(!Qty) Or !Qty = 0, "1", !Qty))  'IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         TxtDiscPC.Text = IIf(IsNull(!DiscPC), 0, !DiscPC)
         TxtDiscPer.Text = IIf(IsNull(!DiscPer), 0, !DiscPer)
         If Val(TxtDiscPC.Text) <> 0 Then
            TxtDiscPer.Text = Round((Val(TxtDiscPC.Text) * 100) / IIf(Val(TxtPrice.Text) = 0, 1, Val(TxtPrice.Text)), 2)
         End If
         If ObjRegistry.ShowStockPriceChecker = True Then
         If ObjRegistry.ShowAllStoreStock = True Then
            vStrSQL = "select isnull(dbo.FunStock('" & TxtCode.Text & "',Null,0,0,0,0,0,0,'" & Date + 1 & "',0),0)"
            With CN.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyloose = .Fields(0).Value
               Else
                  vQtyloose = 0
               End If
            End With
            LblAllStock.Caption = CN.Execute("SELECT dbo.FunGetPack('" & TxtCode.Text & "',(" & vQtyloose & "))").Fields(0).Value
'            With CN.Execute("Select isnull(abbreviation,'') from packings where packingname = '" & CmbPackName.Text & "'")
'               If .RecordCount > 0 Then
'                  LblAllStock.Caption = LblAllStock.Caption & " " & .Fields(0).Value
'               Else
'                  LblAllStock.Caption = LblAllStock.Caption & " "
'               End If
'            End With
            LblAllStock.Caption = LblAllStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose('" & TxtCode.Text & "',(" & vQtyloose & "))").Fields(0).Value
            LblAllStock.Caption = LblAllStock.Caption & " " & "Loose"
            LblAllStock.Visible = True
         Else
            LblAllStock.Visible = False
         End If
         
          If ObjRegistry.ShowSavedStock = True Then
            vStrSQL = "select qtyloose from currentStockStore where Storeid = " & TxtStoreID.Text & " and Productid = '" & TxtPID.Text & "'"
            With CN.Execute(vStrSQL)
               If .RecordCount > 0 Then
                  vQtyloose = .Fields(0).Value
               Else
                  vQtyloose = 0
               End If
            End With
         Else
            vStrSQL = "select isnull(dbo.FunStock('" & TxtCode.Text & "'," & IIf((ObjRegistry.StoreID = ""), 1, ObjRegistry.StoreID) & ",0,0,0,0,0,0,'" & Date & "',0),0)"
            vQtyloose = CN.Execute(vStrSQL).Fields(0).Value
         End If
         LblStock.Caption = CN.Execute("SELECT dbo.FunGetPack('" & TxtCode.Text & "',Floor(" & vQtyloose & "))").Fields(0).Value
'         LblStock.Caption = LblStock.Caption & " " & CmbPackName.Text
         LblStock.Caption = LblStock.Caption & " " & CN.Execute("SELECT dbo.FunGetLoose('" & TxtCode.Text & "',Floor(" & vQtyloose & "))").Fields(0).Value
         LblStock.Caption = LblStock.Caption & " " & "Loose"
         LblStock.Visible = True
         LblStockCaption.Visible = True
         End If
         SubCalculateBody
         FunSelectProduct = True
      Else
         FunSelectProduct = False
         .Close
         TxtCode.Text = ""
         TxtProductName.Text = ""
         TxtProductName2.Caption = ""
         TxtQty.Text = ""
         TxtPrice.Text = ""
         TxtSC.Text = ""
         TxtDiscPC.Text = ""
         TxtDiscPer.Text = ""
         TxtDiscVal.Text = ""
         TxtAmount.Text = ""
      End If
   End With
   TxtCode.Text = ""
   TxtCode.SetFocus
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub SubCalculateBody()
   TxtActualAmount.Text = Val(TxtQty.Text) * (Val(TxtPrice.Text) + Val(TxtSC.Text))
   TxtOriginalPrice.Caption = TxtActualAmount.Text
   TxtDiscVal.Text = Val(TxtQty.Text) * Val(TxtDiscPC.Text)
   TxtAmount.Text = Val(TxtActualAmount.Text) - Val(TxtDiscVal.Text)
   TxtNetAmount.Caption = TxtAmount.Text
   TxtProductName2.Caption = TxtProductName.Text
End Sub

Private Sub TxtNetAmount_Change()
   On Error GoTo ErrorHandler
   If vShowDiscPrice Then
      If Len(TxtNetAmount.Caption) > 5 Then
         TxtNetAmount.FontSize = 144
      ElseIf Len(TxtNetAmount.Caption) > 3 Then
         TxtNetAmount.FontSize = 168
      Else
         TxtNetAmount.FontSize = 192
      End If
   Else
      If Len(TxtNetAmount.Caption) > 5 Then
         TxtNetAmount.FontSize = 110
         TxtOriginalPrice.FontSize = 110
      ElseIf Len(TxtNetAmount.Caption) > 3 Then
         TxtNetAmount.FontSize = 120
         TxtOriginalPrice.FontSize = 120
      Else
         TxtNetAmount.FontSize = 168
         TxtOriginalPrice.FontSize = 168
      End If
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


