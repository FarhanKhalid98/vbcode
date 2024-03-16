VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptPurchaseDateWise 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "RptPurchaseDateWise.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4666
      TabIndex        =   16
      Top             =   3360
      Width           =   2655
      Begin VB.OptionButton RdoSummary 
         BackColor       =   &H80000009&
         Caption         =   "Summary"
         Height          =   255
         Left            =   1335
         TabIndex        =   18
         Top             =   0
         Width           =   1200
      End
      Begin VB.OptionButton RdoDetail 
         BackColor       =   &H80000009&
         Caption         =   "Detail"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3646
      TabIndex        =   12
      Top             =   2760
      Width           =   4695
      Begin VB.OptionButton RdoBoth 
         BackColor       =   &H80000009&
         Caption         =   "Both Invoices"
         Height          =   255
         Left            =   3150
         TabIndex        =   15
         Top             =   0
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton RdoReturn 
         BackColor       =   &H80000009&
         Caption         =   "Return Invioce"
         Height          =   255
         Left            =   1695
         TabIndex        =   14
         Top             =   0
         Width           =   1455
      End
      Begin VB.OptionButton RdoInv 
         BackColor       =   &H80000009&
         Caption         =   "Purchase Invoice"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1695
      End
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6750
      TabIndex        =   23
      Top             =   7200
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Close"
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
      MICON           =   "RptPurchaseDateWise.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3975
      TabIndex        =   21
      Top             =   7200
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Pre&view"
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
      MICON           =   "RptPurchaseDateWise.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5355
      TabIndex        =   22
      Top             =   7200
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Print"
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
      MICON           =   "RptPurchaseDateWise.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   8910
      TabIndex        =   24
      Top             =   1590
      Visible         =   0   'False
      Width           =   1860
      _ExtentX        =   3281
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
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2865
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5325
      Visible         =   0   'False
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
      MICON           =   "RptPurchaseDateWise.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   1845
      TabIndex        =   0
      Top             =   5325
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   3
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
   Begin SITextBox.Txt TxtOrganizatonName 
      Height          =   315
      Left            =   3225
      TabIndex        =   2
      Tag             =   "nc"
      Top             =   5325
      Visible         =   0   'False
      Width           =   3585
      _ExtentX        =   6324
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
   Begin JeweledBut.JeweledButton BtnStore 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2865
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   4620
      Visible         =   0   'False
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
      MICON           =   "RptPurchaseDateWise.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   1845
      TabIndex        =   3
      Top             =   4620
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   3
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
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   3225
      TabIndex        =   5
      Tag             =   "nc"
      Top             =   4635
      Visible         =   0   'False
      Width           =   3585
      _ExtentX        =   6324
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
   Begin SITextBox.Txt TxtCode 
      Height          =   315
      Left            =   3518
      TabIndex        =   9
      Top             =   2085
      Width           =   1020
      _ExtentX        =   1799
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
   End
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   4538
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2085
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
      MICON           =   "RptPurchaseDateWise.frx":0F56
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   4898
      TabIndex        =   11
      Top             =   2085
      Width           =   3585
      _ExtentX        =   6324
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
   Begin JeweledBut.JeweledButton BtnVendor 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2865
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   5985
      Visible         =   0   'False
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
      MICON           =   "RptPurchaseDateWise.frx":0F72
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtPartyID 
      Height          =   315
      Left            =   1845
      TabIndex        =   6
      Top             =   5985
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   3
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
   Begin SITextBox.Txt TxtPartyName 
      Height          =   315
      Left            =   3225
      TabIndex        =   8
      Tag             =   "nc"
      Top             =   5985
      Visible         =   0   'False
      Width           =   3585
      _ExtentX        =   6324
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   4470
      TabIndex        =   19
      Top             =   4230
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpTo 
      Height          =   315
      Left            =   6225
      TabIndex        =   20
      Top             =   4230
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      Height          =   195
      Left            =   4470
      TabIndex        =   36
      Top             =   4005
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      Height          =   195
      Left            =   6240
      TabIndex        =   35
      Top             =   4005
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Name"
      Height          =   195
      Left            =   3225
      TabIndex        =   34
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor ID"
      Height          =   195
      Left            =   1845
      TabIndex        =   33
      Top             =   5760
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   4898
      TabIndex        =   32
      Top             =   1890
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   3518
      TabIndex        =   31
      Top             =   1890
      Width           =   375
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
      Height          =   195
      Left            =   1845
      TabIndex        =   30
      Top             =   5100
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      Height          =   195
      Left            =   3225
      TabIndex        =   29
      Top             =   5100
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Store ID"
      Height          =   195
      Left            =   1845
      TabIndex        =   28
      Top             =   4395
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Store Name"
      Height          =   195
      Left            =   3225
      TabIndex        =   27
      Top             =   4395
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Date Wise"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1935
      TabIndex        =   26
      Top             =   255
      Width           =   2295
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ProductID"
      Height          =   195
      Left            =   8910
      TabIndex        =   25
      Top             =   1395
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove this Row"
      End
   End
End
Attribute VB_Name = "RptPurchaseDateWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim sSql As String

Private Function FunSelectVendor(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchVendor.Show vbModal, Me
        If SchVendor.ParaOutVendorID = "" Then FunSelectVendor = False: Exit Function
        TxtPartyID.Text = SchVendor.ParaOutVendorID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Parties where PartyID=" & Val(TxtPartyID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtPartyName.Text = !PartyName
          FunSelectVendor = True
          .Close
          Exit Function
      Else
          FunSelectVendor = False
          .Close
          TxtPartyID.Text = ""
          TxtPartyName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectOrganizaton(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchOrganization.Show vbModal, Me
        If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganizaton = False: Exit Function
       TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
    End If
    If TxtOrganizationID.Text = "" Then FunSelectOrganizaton = False: Exit Function
    vStrSQL = " Select * FROM Organizations where OrganizationID='" & TxtOrganizationID.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtOrganizatonName.Text = !OrganizationName
          FunSelectOrganizaton = True
          .Close
          Exit Function
      Else
          FunSelectOrganizaton = False
          .Close
         TxtOrganizationID.Text = ""
          TxtOrganizatonName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function


Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtCode.Text = SchProduct.ParaOutID
   End If
    '---------------------------
    If Trim(TxtCode.Text) = "" Then Exit Function
    If Len(TxtCode.Text) <= 5 Then
      TxtCode.Text = Right("00000" + CStr(Val(TxtCode.Text)), 5)
    End If
    If TxtCode.Text = "" Then FunSelectProduct = False: Exit Function
    vStrSQL = " SELECT p.productid, code, ProductName" & vbCrLf _
           + " from Products p left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where p.productid = '" & TxtCode.Text & "' or code='" & TxtCode.Text & "'"
  
   With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductID.Text = !Productid
         TxtProductName.Text = !ProductName
         FunSelectProduct = True
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         MsgBox "Invalid Product ID.", vbOKOnly, "Alert"
         TxtProductID.Text = ""
         TxtCode.Text = ""
         TxtProductName.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
     RdoInv.SetFocus
   Else
      TxtCode.SetFocus
   End If
End Sub

Private Sub BtnGroup_Click()
   If FunSelectOrganizaton(ssButton, False) = True Then
      TxtStoreID.SetFocus
   Else
     TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub BtnVendor_Click()
   If FunSelectVendor(ssButton, False) = True Then
      TxtCode.SetFocus
   Else
      TxtPartyID.SetFocus
   End If
End Sub

Private Sub BtnStore_Click()
If FunSelectStore(ssButton, False) = True Then
     TxtCode.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub TxtCode_Change()
   If ActiveControl.Name <> TxtCode.Name Then Exit Sub
   If TxtProductName.Text <> "" Then
      TxtCode.Text = ""
      TxtProductID.Text = ""
      TxtProductName.Text = ""
   End If
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


Private Sub TxtPartyID_Change()
   If TxtPartyID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtPartyID.Name Then Exit Sub
   If TxtPartyName.Text <> "" Then TxtPartyName.Text = ""
End Sub

Private Sub TxtPartyID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtPartyID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtPartyID.Text = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectVendor(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectVendor(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtOrganizationID_Change()
   If TxtOrganizationID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   If TxtOrganizatonName.Text <> "" Then TxtOrganizatonName.Text = ""
End Sub

Private Sub TxtOrganizationID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtOrganizationID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectOrganizaton(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectOrganizaton(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub



Private Sub TxtStoreID_Change()
If TxtStoreID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   If TxtStoreName.Text <> "" Then TxtStoreName.Text = ""
End Sub

Private Sub TxtStoreID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtStoreID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtStoreID.Text) = "" Then Exit Sub
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


Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
       If RdoDetail.Value = True Then RptReportViewer.Caption = "Purchase Detail (Date Wise)" Else RptReportViewer.Caption = "Purchase Summary (Date Wise)"
       RptReportViewer.Show vbModal
   End If
End Sub

Private Sub BtnPrint_Click()
    If SetReport Then RptReportViewer.Report.PrintOut False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
         Case vbKeyV
            If BtnPreview.Enabled Then BtnPreview_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
        Case TxtOrganizationID.Name: If FunSelectOrganizaton(ssFunctionKey, True) = True Then TxtStoreID.SetFocus
        Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, True) = True Then TxtCode.SetFocus
        Case TxtPartyID.Name: If FunSelectVendor(ssFunctionKey, True) = True Then TxtCode.SetFocus
        Case TxtCode.Name: If FunSelectProduct(ssFunctionKey, True) = True Then RdoInv.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   ShowPicture Me
   SetWindowText Me.hWnd, "Purchase Detail (Date Wise)"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   Dim frmObj As Object
   For Each frmObj In Forms
       Set frmObj = Nothing
   Next
   'Set RsReport = Nothing
   Set RptPurchaseDateWise = Nothing
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Function SetReport() As Boolean
    On Error GoTo ErrorHandler
    SetReport = False
    Me.MousePointer = vbHourglass
    Dim RsReport As New ADODB.Recordset
    'Set RsReport = CN.Execute("EXEC ProdRptProductStockSummary '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtStoreID.Text) = "", "Null", "'" & TxtStoreID.Text) & "'," & IIf(Trim(TxtProductID.Text) = "", "Null", "'" & TxtProductID.Text & "'") & "," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" &TxtOrganizationID.Text & "'") & "," & IIf(Trim(TxtPartyID.Text) = "", "Null", "'" & TxtPartyID.Text & "'") & "," & IIf(Trim(TxtSubGroupID.Text) = "", "Null", "'" & TxtSubGroupID.Text & "'"))
'    sSQL = "EXEC ProdRptSaleRegister '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", "'" & TxtStoreID.Text & "'") & "," & IIf(Trim(TxtPartyID.Text) = "", "Null", "'" & TxtPartyID.Text & "'") & "," & IIf(Trim(TxtProductID.Text) = "", "Null", "'" & TxtProductID.Text & "'") & "," & IIf(RdoBoth.Value = True, "Null", IIf(RdoInv.Value = True, 0, 1))
    If RdoDetail.Value = True Then
        Set RsReport = CN.Execute("EXEC ProdRptPurchaseRegister '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", "'" & TxtStoreID.Text & "'") & "," & IIf(Trim(TxtPartyID.Text) = "", "Null", "'" & TxtPartyID.Text & "'") & "," & IIf(Trim(TxtProductID.Text) = "", "Null", "'" & TxtProductID.Text & "'") & "," & IIf(RdoBoth.Value = True, "Null", IIf(RdoInv.Value = True, 0, 1)))
    Else
        Set RsReport = CN.Execute("EXEC ProdRptPurchaseRegisterSummary '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'," & IIf(Trim(TxtOrganizationID.Text) = "", "Null", "'" & TxtOrganizationID.Text & "'") & "," & IIf(Trim(TxtStoreID.Text) = "", "Null", "'" & TxtStoreID.Text & "'") & "," & IIf(Trim(TxtPartyID.Text) = "", "Null", "'" & TxtPartyID.Text & "'") & "," & IIf(RdoBoth.Value = True, "Null", IIf(RdoInv.Value = True, 0, 1)))
    End If
    If RdoDetail.Value = True Then Set RptReportViewer.Report = New CrptPurchaseDateWise Else Set RptReportViewer.Report = New CrptPurchaseDateSummary
    If RsReport.BOF Then
        MsgBox "No record exists.", vbInformation, Me.Caption
        Me.MousePointer = vbDefault
        Exit Function
    End If
    If RdoDetail.Value = True Then RptReportViewer.Report.ReportTitle = "Purchase Detail (Date Wise)" Else RptReportViewer.Report.ReportTitle = "Purchase Summary (Date Wise)"
    RptReportViewer.Report.Database.SetDataSource RsReport
    RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
    RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
    RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
    RptReportViewer.Report.ParameterFields(4).AddCurrentValue " Date From :" & Format(DtpFrom.DateValue, "dd/MM/yyyy") & " To : " & Format(DtpTo.DateValue, "dd/MM/yyyy")
    RptReportViewer.Report.ParameterFields(5).AddCurrentValue ObjRegistry.DevelopedBy
    RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
    RptReportViewer.Report.PaperOrientation = crPortrait
       
    RptReportViewer.Report.SelectPrinter "Dummy Driver", "Ding Dong", "LPT1"
    'RptReportViewer.Report.PaperOrientation = crLandscape
    SetReport = True
    Me.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Function FunSelectStore(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchStore.Show vbModal, Me
        If SchStore.ParaOutStoreID = "" Then FunSelectStore = False: Exit Function
        TxtStoreID.Text = SchStore.ParaOutStoreID
    End If
    '---------------------------
    If Trim(TxtStoreID.Text) = "" Then Exit Function
    If Len(TxtStoreID.Text) <= 3 Then
      TxtStoreID.Text = Right("000" + CStr(Val(TxtStoreID.Text)), 3)
    End If
    If TxtStoreID.Text = "" Then FunSelectStore = False: Exit Function
    vStrSQL = " Select StoreName FROM Stores where StoreID='" & TxtStoreID.Text & "'"
    With CN.Execute(vStrSQL)
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

