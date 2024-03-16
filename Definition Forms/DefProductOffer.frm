VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form DefProductOffer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkFixedRebate 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Fixed"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10470
      TabIndex        =   8
      Tag             =   "NC"
      Top             =   1350
      Width           =   795
   End
   Begin VB.TextBox TxtFilterID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2265
      MaxLength       =   5
      TabIndex        =   10
      Top             =   4020
      Width           =   900
   End
   Begin VB.TextBox TxtFilter 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3165
      MaxLength       =   30
      TabIndex        =   13
      Text            =   "Search"
      Top             =   4020
      Width           =   2865
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4170
      Left            =   1928
      TabIndex        =   15
      Top             =   4350
      Width           =   11505
      ScrollBars      =   2
      _Version        =   196616
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "DefProductOffer.frx":0000
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
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
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   8
      Columns(0).Width=   1588
      Columns(0).Caption=   "Product ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4948
      Columns(1).Caption=   "Product Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1323
      Columns(2).Caption=   "Qty"
      Columns(2).Name =   "Qty"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1588
      Columns(3).Caption=   "Product ID"
      Columns(3).Name =   "ProductOfferID"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   4948
      Columns(4).Caption=   "Product Name"
      Columns(4).Name =   "ProductOfferName"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1323
      Columns(5).Caption=   "Qty"
      Columns(5).Name =   "QtyOffer"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1323
      Columns(6).Caption=   "Rebate"
      Columns(6).Name =   "Rebate"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   2196
      Columns(7).Caption=   "Fixed Rebate"
      Columns(7).Name =   "FixedRebate"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   11
      Columns(7).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20294
      _ExtentY        =   7355
      _StockProps     =   79
      BackColor       =   15724527
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   5033
      TabIndex        =   16
      Top             =   9105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "New"
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
      MICON           =   "DefProductOffer.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   6373
      TabIndex        =   17
      Top             =   9105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Change"
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
      MICON           =   "DefProductOffer.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   7713
      TabIndex        =   18
      Top             =   9105
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
      MICON           =   "DefProductOffer.frx":0054
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6390
      TabIndex        =   11
      Top             =   3090
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
      MICON           =   "DefProductOffer.frx":0070
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7710
      TabIndex        =   12
      Top             =   3090
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
      MICON           =   "DefProductOffer.frx":008C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9053
      TabIndex        =   14
      Top             =   9105
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
      MICON           =   "DefProductOffer.frx":00A8
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnProduct 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3270
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2520
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
      MICON           =   "DefProductOffer.frx":00C4
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   2550
      TabIndex        =   0
      Top             =   2520
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
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   3630
      TabIndex        =   2
      Tag             =   "nc"
      Top             =   2520
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
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   6615
      TabIndex        =   3
      Top             =   2520
      Width           =   600
      _ExtentX        =   1058
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
   Begin JeweledBut.JeweledButton BtnProductOffer 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8145
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   2520
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
      MICON           =   "DefProductOffer.frx":00E0
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductOfferID 
      Height          =   315
      Left            =   7425
      TabIndex        =   4
      Top             =   2520
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
   Begin SITextBox.Txt TxtProductOfferName 
      Height          =   315
      Left            =   8505
      TabIndex        =   6
      Tag             =   "nc"
      Top             =   2520
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
   Begin SITextBox.Txt TxtQtyOffer 
      Height          =   315
      Left            =   11490
      TabIndex        =   7
      Top             =   2520
      Width           =   600
      _ExtentX        =   1058
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
   Begin SITextBox.Txt TxtRebate 
      Height          =   315
      Left            =   10920
      TabIndex        =   9
      Top             =   1845
      Width           =   600
      _ExtentX        =   1058
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
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Offers"
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
      TabIndex        =   30
      Top             =   270
      Width           =   1905
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rebate / Discount"
      Height          =   195
      Left            =   10440
      TabIndex        =   29
      Top             =   1605
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      Height          =   195
      Left            =   2265
      TabIndex        =   28
      Top             =   3810
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Offer"
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
      Left            =   8865
      TabIndex        =   27
      Top             =   1920
      Width           =   990
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Product"
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
      Left            =   4065
      TabIndex        =   26
      Top             =   1920
      Width           =   1305
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   11520
      TabIndex        =   25
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   8505
      TabIndex        =   24
      Top             =   2280
      Width           =   1020
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      Height          =   195
      Left            =   7425
      TabIndex        =   23
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   6615
      TabIndex        =   22
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3630
      TabIndex        =   21
      Top             =   2280
      Width           =   1020
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      Height          =   195
      Left            =   2550
      TabIndex        =   20
      Top             =   2280
      Width           =   765
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3165
      TabIndex        =   19
      Top             =   3810
      Width           =   1020
   End
End
Attribute VB_Name = "DefProductOffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim RsOffer As New ADODB.Recordset
Dim vMode As FormMode, vProductName As String
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.

Private Sub BtnClear_Click()
  FormStatus = SelectionMode
  TxtFilterID.SetFocus
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      TxtQty.SetFocus
   Else
      TxtProductID.SetFocus
   End If
End Sub

Private Sub BtnProductOffer_Click()
 If FunSelectProductOffer(ssButton, True) = True Then
      TxtQtyOffer.SetFocus
   Else
      TxtProductOfferID.SetFocus
   End If
End Sub


Private Sub ChkFixedRebate_Click()
If ChkFixedRebate.Visible = False Then Exit Sub
If Me.ActiveControl.Name <> ChkFixedRebate.Name Then Exit Sub
If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
    If KeyCode = vbKeyReturn Then
      If ActiveControl.Name = Grid.Name Then Call Grid_DblClick: Exit Sub
      keybd_event 9, 1, 1, 1
      KeyCode = 0
    End If
    If Shift = vbCtrlMask Then
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
            Case vbKeyO
                If BtnOpen.Enabled Then BtnOpen_Click
                KeyCode = 0
            Case vbKeyR
                If BtnDelete.Enabled Then BtnDelete_Click
                KeyCode = 0
        End Select
    ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtProductID.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtQty.SetFocus
         Case TxtProductOfferID.Name: If FunSelectProductOffer(ssFunctionKey, True) = True Then TxtQtyOffer.SetFocus
      End Select
    End If
    Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniProductOffer", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
  Dim vtbl As String
  Rs.Filter = 0
  Rs.Filter = "ProductID = '" & Grid.Columns("id").Text & "'"
  If Rs.RecordCount > 0 Then
  If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
    vtbl = Common.ChildDataExists("ProductOffers", "ProductId='" & Rs!Productid & "'", "")
    If vtbl <> "" Then
      MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
      Exit Sub
    End If
    Rs.Delete
    Rs.Filter = 0
    If Rs.RecordCount = 0 Then FormStatus = NewMode: Exit Sub
    Rs.MoveNext
    Grid.MoveNext
    If Rs.EOF Then Rs.MoveLast
    Call Form_Load
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnNew_Click()
  FormStatus = NewMode
  TxtProductID.SetFocus
End Sub

Private Sub BtnOpen_Click()
  On Error GoTo ErrorHandler
  If Rs.RecordCount > 0 Then
    If Rs.BOF = False And Rs.EOF = False Then
      FormStatus = OpenMode
    End If
  End If
  TxtQty.SetFocus
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
'   If FunValidation = False Then Exit Sub
   If vIsNewRecord Then
      If cn.Execute("Select ProductID from ProductOffers where ProductID = '" & TxtProductID.Text & "'").RecordCount > 0 Then
        MsgBox "This product ID already exists : ", vbInformation, "Alert"
        TxtProductID.SetFocus
        Exit Sub
      End If
      Rs.AddNew
      Rs!Productid = TxtProductID.Text
      Rs!Qty = TxtQty.Text
      Rs!ProductOfferID = IIf(Trim(TxtProductOfferID.Text) = "", Null, TxtProductOfferID.Text)
      Rs!QtyOffer = Val(TxtQtyOffer.Text)
      Rs!Rebate = Val(TxtRebate.Text)
      Rs!FixedRebate = ChkFixedRebate.Value
   Else
      Rs.Filter = "Productid ='" & TxtProductID.Text & "'"
      Rs!Qty = TxtQty.Text
      Rs!ProductOfferID = IIf(Trim(TxtProductOfferID.Text) = "", Null, TxtProductOfferID.Text)
      Rs!QtyOffer = Val(TxtQtyOffer.Text)
      Rs!FixedRebate = ChkFixedRebate.Value
     Rs!Rebate = Val(TxtRebate.Text)
    End If
    Rs.Filter = 0
   Rs.Update
   FormStatus = NewMode
   Call Form_Load
   TxtProductID.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
'  If vIsNewRecord Then
'    If Trim(TxtProductID.Text) = "" Then
'      MsgBox "Please specify a Group ID", vbExclamation, "Alert"
'      If TxtProductID.Enabled And TxtProductID.Visible Then TxtProductID.SetFocus
'      Exit Function
'    End If
'    If Len(Trim(TxtProductID.Text)) < 3 Then
'      MsgBox "The Group ID must be three characters long", vbExclamation, "Alert"
'      TxtProductID.Text = Right("000" + CStr(Val(TxtProductID.Text)), 3)
'      If TxtProductID.Enabled And TxtProductID.Visible Then TxtProductID.SetFocus
'      Exit Function
'    End If
'    If CN.Execute("Select count(*) from Groups where Groupid = '" & TxtProductID.Text & "'").Fields(0) > 0 Then
'        MsgBox "This Group ID already exists. The Group ID must be unique", vbExclamation, "Alert"
'        TxtProductID.Text = FunGetMaxID
'        If TxtProductID.Enabled And TxtProductID.Visible Then TxtProductID.SetFocus
'        Exit Function
'    End If
'    Select Case Asc(UCase(Left(TxtProductID.Text, 1)))
'      Case 65 To 90
'      Case 48 To 57
'      Case Else
'        MsgBox "The Group ID must contain numeric/alphabetical characters only", vbExclamation, "Alert"
'        If TxtProductID.Enabled And TxtProductID.Visible Then TxtProductID.SetFocus
'        Exit Function
'    End Select
'    Select Case Asc(UCase(Right(TxtProductID.Text, 1)))
'      Case 65 To 90
'      Case 48 To 57
'      Case Else
'        MsgBox "The Group ID must contain numeric/alphabetical characters only", vbExclamation, "Alert"
'        If TxtProductID.Enabled And TxtProductID.Visible Then TxtProductID.SetFocus
'        Exit Function
'    End Select
'  End If
'  If Trim(TxtProductName.Text) = "" Then
'    MsgBox "Please specify a Group name", vbExclamation, "Alert"
'    If TxtProductName.Enabled And TxtProductName.Visible Then TxtProductName.SetFocus
'    Exit Function
'  End If
'  'All Ok, now validation is success
'  FunValidation = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Product Offers"
   Me.MousePointer = vbHourglass
   Set Rs = New ADODB.Recordset
   Rs.Open "Select ProductID, Qty, ProductOfferID, QtyOffer, Rebate, FixedRebate  FROM ProductOffers", cn, adOpenDynamic, adLockOptimistic
        
   Grid.Columns("ID").DataField = "ProductID"
   Grid.Columns("Name").DataField = "ProductName"
   Grid.Columns("ProductOfferName").DataField = "ProductOfferName"
   Grid.Columns("Qty").DataField = "Qty"
   Grid.Columns("ProductOfferID").DataField = "ProductOfferID"
   Grid.Columns("QtyOffer").DataField = "QtyOffer"
   Grid.Columns("Rebate").DataField = "Rebate"
   Grid.Columns("FixedRebate").DataField = "FixedRebate"
   LoadGrid
   FormStatus = NewMode
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
    Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   'VStrSQL = "SELECT p.productid, Code, productname, retailprice,retailprice-discpc as DiscPrice FROM products p inner join (select * from ProductBarcodes where len(code)=7)b on p.productid = b.productid where productname like '%" & TxtProductName.Text & "%'" & vOrder & vDirection
   Dim VStrSQL As String
   
'   VStrSQL = "Select PO.*, P.ProductName, PON.ProductName as ProductOfferName " & vbCrLf _
            + " from ProductOffers PO inner Join Products p on p.productid = po.productid " & vbCrLf _
            + " Left Outer Join Products PON on PON.ProductID = PO.ProductOfferID " & vbCrLf _
            + " left outer join ProductBarcodes b on PON.ProductID = b.productid" & vbCrLf _
            + " where (p.productid like '%" & TxtFilterID.Text & "%' or code = '" & TxtFilterID.Text & "')" & vProductName
            
   VStrSQL = "Select PO.*, P.ProductName, PON.ProductName as ProductOfferName " & vbCrLf _
            + " from ProductOffers PO inner Join Products p on p.productid = po.productid " & vbCrLf _
            + " Left Outer Join Products PON on PON.ProductID = PO.ProductOfferID " & vbCrLf _
            + " where (p.productid like '%" & TxtFilterID.Text & "%')" & vProductName
   
   If RsOffer.State = adStateOpen Then RsOffer.Close
   RsOffer.CursorLocation = adUseClient
   RsOffer.Open VStrSQL, cn, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = RsOffer
   'Grid.DataMode
   'Flag = False
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
         BtnNew.Enabled = False
         BtnOpen.Enabled = False
         BtnDelete.Enabled = False
         BtnSave.Enabled = False
         BtnClear.Enabled = True
         BtnProductOffer.Enabled = True
         TxtQty.Enabled = True
         TxtProductOfferID.Enabled = True
         TxtQtyOffer.Enabled = True
         TxtRebate.Enabled = True
         TxtProductID.Text = ""
         TxtProductName.Text = ""
         TxtQty.Text = ""
         TxtProductOfferID.Text = ""
         TxtProductOfferName.Text = ""
         TxtQtyOffer.Text = ""
         TxtRebate.Text = ""
         ChkFixedRebate.Enabled = True
         TxtFilterID.Text = ""
         TxtFilter.Text = ""
         TxtFilter.Enabled = False
         TxtFilter.BackColor = &HE0E0E0
         TxtFilterID.Enabled = False
         TxtFilterID.BackColor = &HE0E0E0
'         TxtProductID.Text = FunGetMaxID()
         TxtProductID.Enabled = True
         BtnProduct.Enabled = True
         TxtProductID.BackColor = vbWhite
         Grid.Enabled = False
         'If TxtProductID.Enabled And TxtProductID.Visible Then TxtProductID.SetFocus
         vIsNewRecord = True
     Case Is = OpenMode
         TxtFilter.Text = ""
         TxtFilterID.Text = ""
         BtnNew.Enabled = False
         BtnOpen.Enabled = False
         BtnDelete.Enabled = False
         BtnClear.Enabled = True
         Grid.Enabled = False
         TxtFilter.Enabled = False
         TxtFilter.BackColor = &HE0E0E0
         TxtFilterID.Enabled = False
         TxtFilterID.BackColor = &HE0E0E0
         TxtProductID.Enabled = False
         BtnProduct.Enabled = False
         TxtProductID.BackColor = &HE0E0E0
         TxtQty.Enabled = True
         TxtProductOfferID.Enabled = True
         BtnProductOffer.Enabled = True
         TxtQtyOffer.Enabled = True
         TxtRebate.Enabled = True
         ChkFixedRebate.Enabled = True
         
         vIsNewRecord = False
     Case Is = ChangeMode
         BtnSave.Enabled = True
     Case Is = SelectionMode
         Grid.Enabled = True
         TxtFilter.Text = ""
         TxtFilter.Enabled = True
         TxtFilter.BackColor = vbWhite
         TxtFilterID.Enabled = True
         TxtFilterID.BackColor = vbWhite
         BtnNew.Enabled = True
         BtnOpen.Enabled = True
         BtnDelete.Enabled = True
         BtnSave.Enabled = False
         BtnClear.Enabled = False
         TxtProductID.Enabled = False
         BtnProduct.Enabled = False
         TxtQty.Enabled = False
         TxtProductOfferID.Enabled = False
         BtnProductOffer.Enabled = False
         TxtQtyOffer.Enabled = False
         TxtRebate.Enabled = False
         ChkFixedRebate.Enabled = False
         TxtProductID.BackColor = &HE0E0E0
         Call Grid_RowColChange(0, 0)
         Grid.SetFocus
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

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
      Set RsOffer = Nothing
      Set DefProductOffer = Nothing
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKeyA To vbKeyZ, vbKey0 To vbKey9, Asc("a") To Asc("z")
      TxtFilter.Text = Chr(KeyAscii): TxtFilter.SelStart = Len(TxtFilter.Text):  TxtFilter.SetFocus
   End Select
   
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo ErrorHandler
   If RsOffer.RecordCount > 0 And Grid.Enabled Then
      TxtProductID.Text = Grid.Columns("ID").Text
      TxtProductName.Text = Grid.Columns("Name").Text
      TxtQty.Text = Grid.Columns("Qty").Text
      
      TxtProductOfferID.Text = Grid.Columns("ProductOfferID").Text
      TxtProductOfferName.Text = Grid.Columns("ProductOfferName").Text
      TxtQtyOffer.Text = Grid.Columns("QtyOffer").Text
      
      TxtRebate.Text = Grid.Columns("Rebate").Text
      ChkFixedRebate.Value = Abs(Grid.Columns("FixedRebate").Value)
      
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
   If Grid.Rows > 0 And BtnOpen.Enabled Then BtnOpen_Click
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtFilter_Change()
   On Error GoTo ErrorHandler
   Dim vWords
   vWords = Split(TxtFilter.Text, " ")
   vProductName = ""
   For i = 0 To UBound(vWords)
       vProductName = vProductName & " and p.Productname like '%" & Replace(vWords(i), "'", "''") & "%'"
   Next
'   vProductName = " and Productname like '%" & Replace(TxtProductName.Text, "'", "''") & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
'
'   On Error GoTo ErrorHandler
'   If Trim(TxtFilter.Text) = "" Then Grid.MoveFirst: Exit Sub
'   RsOffer.Find "ProductName like '%" & TxtFilter.Text & "%'", , adSearchForward, 1
'   If RsOffer.EOF Then Grid.MoveLast
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
End Sub

Private Sub TxtFilterID_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtFilterID.Text) = "" Then Grid.MoveFirst: Exit Sub
   RsOffer.Find "ProductID like '" & TxtFilterID.Text & "%'", , adSearchForward, 1
   If RsOffer.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtProductID_Change()
   If TxtProductID.Enabled = False Then Exit Sub
   If Me.ActiveControl.Name <> TxtProductID.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtProductID.Text = ""
      TxtProductName.Text = ""
      TxtQty.Text = ""
   End If
    If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub TxtProductID_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler
    If Me.ActiveControl.Name <> TxtProductID.Name Then Exit Sub
    Dim vTemp As Boolean
    vTemp = Not FunSelectProduct(ssValidate, True)
    If vTemp = True Then
        vTemp = Not FunSelectProduct(ssButton, False)
    End If
    Cancel = vTemp
Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub TxtProductName_Change()
   If TxtProductName.Enabled = True Then FormStatus = ChangeMode
End Sub

Private Sub TxtProductName_LostFocus()
   TxtProductName.Text = StrConv(TxtProductName.Text, vbProperCase)
End Sub

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtProductID.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtProductID.Text) = "" Then Exit Function
    If Len(TxtProductID.Text) <= 5 Then
      TxtProductID.Text = Right("00000" + CStr(Val(TxtProductID.Text)), 5)
    End If
    If TxtProductID.Text = "" Then FunSelectProduct = False: Exit Function
    
   VStrSQL = " SELECT p.ProductID, Code, Qty, ProductName" & vbCrLf _
         + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
         + " where (p.productid = '" & TxtProductID.Text & "' or Code = '" & TxtProductID.Text & "')" & " and isLocked = 0 "

  With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         TxtQty.Text = IIf(IsNull(!Qty) Or !Qty = 0, "1", !Qty)  'IIf(Val(TxtQty.Text) = 0, 1, TxtQty.Text)
         FunSelectProduct = True
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtProductID.Text = ""
         TxtProductName.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtProductOfferID_Change()
   If TxtProductOfferID.Enabled = False Then Exit Sub
   If Me.ActiveControl.Name <> TxtProductOfferID.Name Then Exit Sub
   If TxtProductOfferName.Text <> "" Then
      TxtProductOfferID.Text = ""
      TxtProductOfferName.Text = ""
      TxtQtyOffer.Text = ""
   End If
   TxtRebate.Text = ""
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub TxtProductOfferID_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtProductOfferID.Text) = "" Then Exit Sub
   vTemp = Not FunSelectProductOffer(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectProductOffer(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectProductOffer(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProductOffer = False: Exit Function
      TxtProductOfferID.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtProductOfferID.Text) = "" Then Exit Function
    If Len(TxtProductOfferID.Text) <= 5 Then
      TxtProductOfferID.Text = Right("00000" + CStr(Val(TxtProductOfferID.Text)), 5)
    End If
    If TxtProductOfferID.Text = "" Then FunSelectProductOffer = False: Exit Function
    VStrSQL = " SELECT p.Productid, ProductName" & vbCrLf _
           + " from Products p" & vbCrLf _
           + " where p.productid = '" & TxtProductOfferID.Text & "'"
  
   With cn.Execute(VStrSQL)
      If .RecordCount > 0 Then
         TxtProductOfferID.Text = !Productid
         TxtProductOfferName.Text = !ProductName
         FunSelectProductOffer = True
         .Close
         Exit Function
      Else
         FunSelectProductOffer = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtProductOfferID.Text = ""
         TxtProductOfferName.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtQty_Change()
    If Me.ActiveControl.Name <> TxtQty.Name Then Exit Sub
    If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub TxtQtyOffer_Change()
    If Me.ActiveControl.Name <> TxtQtyOffer.Name Then Exit Sub
    If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub TxtRebate_Change()
    If Me.ActiveControl.Name <> TxtRebate.Name Then Exit Sub
    TxtProductOfferID.Text = ""
    TxtProductOfferName.Text = ""
    TxtQtyOffer.Text = ""
    If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub
