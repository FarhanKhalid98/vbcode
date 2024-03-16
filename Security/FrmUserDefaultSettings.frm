VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmUserDefaultSettings 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmUserDefaultSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkShowStock 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Stock"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6885
      TabIndex        =   41
      Top             =   2610
      Width           =   3210
   End
   Begin VB.CheckBox ChkAllowDiscount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Allow Discount"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6885
      TabIndex        =   40
      Top             =   3015
      Width           =   3210
   End
   Begin VB.TextBox TxtNoofPurPrints 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   4500
      MaxLength       =   100
      TabIndex        =   38
      Top             =   6750
      Width           =   885
   End
   Begin VB.CheckBox ChkPurchaseRePrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purchase Re Print"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10710
      TabIndex        =   37
      Top             =   5130
      Width           =   3210
   End
   Begin VB.CheckBox ChkSaleRePrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sale Re Print"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10710
      TabIndex        =   36
      Top             =   4710
      Width           =   3210
   End
   Begin VB.CheckBox ChkChangePriceFormOpenAsLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Price Form Open as login When Margin below  0"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6930
      TabIndex        =   33
      Top             =   5805
      Width           =   4380
   End
   Begin VB.CheckBox ChkSalePriceMustBeLessThanPurchase 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sale Price must be Less than Purchase"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6885
      TabIndex        =   32
      Top             =   5265
      Width           =   3210
   End
   Begin VB.CheckBox ChkShowSumInSearchSaleInvoice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Sum in Serach Sale Invoice"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6885
      TabIndex        =   31
      Top             =   4815
      Width           =   3210
   End
   Begin VB.CheckBox ChkOpenForm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Allow Open Form"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10665
      TabIndex        =   30
      Top             =   4260
      Width           =   3210
   End
   Begin VB.CheckBox ChkChangeDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Allow Change Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10665
      TabIndex        =   29
      Top             =   3810
      Width           =   3210
   End
   Begin VB.CheckBox ChkShowPurchasePriceInInvoice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Purchase Price In Invoice"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6885
      TabIndex        =   28
      Top             =   3405
      Width           =   3210
   End
   Begin VB.CheckBox ChkNotEditingAfterPrinting 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Not Editing After Printing"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10665
      TabIndex        =   27
      Top             =   3405
      Width           =   3210
   End
   Begin VB.CheckBox ChkDisableCreditSale 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Disable Credit Sale"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6885
      TabIndex        =   26
      Top             =   4260
      Width           =   3210
   End
   Begin VB.CheckBox ChkCreditSale 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Set Credit on sale as Default"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6885
      TabIndex        =   25
      Top             =   3810
      Width           =   3210
   End
   Begin VB.TextBox TxtCommPort 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   8220
      MaxLength       =   100
      TabIndex        =   23
      Top             =   7140
      Width           =   885
   End
   Begin VB.CheckBox ChkShowCustomerPoleDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Customer Pole Display"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7005
      TabIndex        =   22
      Top             =   6555
      Width           =   3210
   End
   Begin VB.CheckBox ChkChangePrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Allow Change Price in Sale POS"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4125
      TabIndex        =   21
      Top             =   7275
      Visible         =   0   'False
      Width           =   3210
   End
   Begin VB.TextBox TxtNoofPrints 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   4500
      MaxLength       =   100
      TabIndex        =   3
      Top             =   6315
      Width           =   885
   End
   Begin VB.ComboBox CmbUsers 
      Height          =   315
      ItemData        =   "FrmUserDefaultSettings.frx":0ECA
      Left            =   2850
      List            =   "FrmUserDefaultSettings.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2745
      Width           =   2730
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7575
      TabIndex        =   7
      Top             =   8205
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
      MICON           =   "FrmUserDefaultSettings.frx":0ECE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6255
      TabIndex        =   4
      Top             =   8205
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
      MICON           =   "FrmUserDefaultSettings.frx":0EEA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3615
      TabIndex        =   6
      Top             =   8205
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
      MICON           =   "FrmUserDefaultSettings.frx":0F06
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8895
      TabIndex        =   8
      Top             =   8205
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
      MICON           =   "FrmUserDefaultSettings.frx":0F22
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4935
      TabIndex        =   5
      Top             =   8205
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
      MICON           =   "FrmUserDefaultSettings.frx":0F3E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtStoreID 
      Height          =   315
      Left            =   2385
      TabIndex        =   1
      Top             =   4140
      Width           =   795
      _ExtentX        =   1402
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtStoreName 
      Height          =   315
      Left            =   3540
      TabIndex        =   10
      Top             =   4140
      Width           =   2415
      _ExtentX        =   4260
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
      Left            =   3180
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4140
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
      MICON           =   "FrmUserDefaultSettings.frx":0F5A
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   2370
      TabIndex        =   2
      Top             =   4785
      Width           =   795
      _ExtentX        =   1402
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
      Mandatory       =   1
   End
   Begin SITextBox.Txt TxtOrganizationName 
      Height          =   315
      Left            =   3525
      TabIndex        =   16
      Top             =   4785
      Width           =   2415
      _ExtentX        =   4260
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
   Begin JeweledBut.JeweledButton BtnOrganization 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   3165
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4785
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
      MICON           =   "FrmUserDefaultSettings.frx":0F76
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtAllowMaximmDiscPer 
      Height          =   315
      Left            =   4500
      TabIndex        =   35
      Top             =   5895
      Width           =   885
      _ExtentX        =   1561
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
      DecimalPoint    =   5
      IntegralPoint   =   2
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No of Prints in Purchase Invoice"
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
      Left            =   1665
      TabIndex        =   39
      Top             =   6795
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allow Maximm DiscPer"
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
      Left            =   3735
      TabIndex        =   34
      Top             =   5625
      Width           =   1905
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comm Port"
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
      Left            =   8175
      TabIndex        =   24
      Top             =   6915
      Width           =   915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No of Prints in Sale Invoice"
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
      Left            =   1980
      TabIndex        =   20
      Top             =   6390
      Width           =   2355
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
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
      Left            =   4125
      TabIndex        =   19
      Top             =   4575
      Width           =   1620
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
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
      Left            =   2370
      TabIndex        =   18
      Top             =   4575
      Width           =   1335
   End
   Begin VB.Label LblStoreID 
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
      Left            =   2385
      TabIndex        =   15
      Top             =   3915
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Box Setting"
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
      Left            =   2250
      TabIndex        =   14
      Top             =   3420
      Width           =   3960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2070
      Left            =   2205
      Top             =   3285
      Width           =   4065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   147
      X2              =   417
      Y1              =   252
      Y2              =   252
   End
   Begin VB.Label Label2 
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
      Left            =   3555
      TabIndex        =   13
      Top             =   3915
      Width           =   1005
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
      Left            =   2865
      TabIndex        =   12
      Top             =   2460
      Width           =   630
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Default Settings"
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
      TabIndex        =   9
      Top             =   270
      Width           =   2910
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11610
      Top             =   60
      Width           =   375
   End
End
Attribute VB_Name = "FrmUserDefaultSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Public RsBody As New ADODB.Recordset
Dim vMode As FormMode
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.
Dim vid As String
Dim sSql As String, vStrSQL As String, vCounter As Integer

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
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
   Dim vTbl As String
'   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
'      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
'      Exit Sub
'   End If
   If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
   CN.BeginTrans
      CN.Execute "DELETE FROM UserRegistry where UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   On Error GoTo ErrorHandler
   SchUserSettings.Show vbModal, Me
   If SchUserSettings.ParaOutUserNo <> 0 Then GetUserSettings
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetUserSettings()
   On Error GoTo ErrorHandler
   
   sSql = "Select UserName, isnull(SaleRePrint,0) SaleRePrint , isnull(PurchaseRePrint,0)  PurchaseRePrint,  h.StoreID, isnull(SaleRePrint,0) SaleRePrint , isnull(PurchaseRePrint,0)  PurchaseRePrint, StoreName, h.OrganizationID, OrganizationName, isnull(NoofPrints,0) as NoofPrints, isnull(NoofPurPrints,0) as NoofPurPrints, isnull(ChangePrice,0) as ChangePrice," & _
   " isnull(CustomerPoleDisplay,0) as CustomerPoleDisplay, ChangePriceFormOpenAsLogin, AllowMaximmDiscPer, CommPort, CreditSale, DisableCreditSale, ShowPurchasePriceInInvoice, ShowSumInSearchSaleInvoice, " & _
   " SalePriceMustBeLessThanPurchase, NotEditingAfterPrinting, ChangeDate, AllowDiscount, ShowStock, OpenForm " & _
   " from UserRegistry h inner join users u on h.UserNo = u.Userno " & _
   " left outer Join Stores s on s.StoreID = H.StoreID " & _
   " left outer Join Organizations o on o.OrganizationID = H.OrganizationID " & _
   " where h.UserNo = " & SchUserSettings.ParaOutUserNo
   
   With CN.Execute(sSql)
      If Not .BOF Then
          TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
          TxtStoreName.Text = IIf(IsNull(!StoreName), "", !StoreName)
          TxtOrganizationID.Text = IIf(IsNull(!OrganizationID), "", !OrganizationID)
          TxtOrganizationName.Text = IIf(IsNull(!OrganizationName), "", !OrganizationName)
          CmbUsers.Text = !UserName
          TxtNoofPrints.Text = !NoofPrints
          TxtNoofPurPrints.Text = !NoofPurPrints
          TxtAllowMaximmDiscPer.Text = !AllowMaximmDiscPer
          ChkChangePrice.Value = Abs(!ChangePrice)
          ChkShowCustomerPoleDisplay.Value = Abs(!CustomerPoleDisplay)
          ChkCreditSale.Value = Abs(!CreditSale)
          ChkDisableCreditSale.Value = Abs(!DisableCreditSale)
          ChkShowPurchasePriceInInvoice.Value = Abs(!ShowPurchasePriceInInvoice)
          ChkShowSumInSearchSaleInvoice.Value = Abs(!ShowSumInSearchSaleInvoice)
          ChkSalePriceMustBeLessThanPurchase.Value = Abs(!SalePriceMustBeLessThanPurchase)
          ChkNotEditingAfterPrinting.Value = Abs(!NotEditingAfterPrinting)
          ChkChangePriceFormOpenAsLogin.Value = Abs(!ChangePriceFormOpenAsLogin)
          ChkChangeDate.Value = Abs(!ChangeDate)
          ChkOpenForm.Value = Abs(!OpenForm)
          ChkSaleRePrint.Value = Abs(!SaleRePrint)
          ChkAllowDiscount.Value = Abs(!AllowDiscount)
          ChkShowStock.Value = Abs(!ShowStock)
          ChkPurchaseRePrint.Value = Abs(!PurchaseRePrint)
          TxtCommPort.Text = IIf(IsNull(!CommPort), "", !CommPort)
          TxtAllowMaximmDiscPer.Text = IIf(IsNull(!AllowMaximmDiscPer), 0, !AllowMaximmDiscPer)
      End If
      .Close
   End With
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnStore_Click()
   If FunSelectStore(ssButton, False) = True Then
      If BtnSave.Enabled = True Then BtnSave.SetFocus
   Else
      TxtStoreID.SetFocus
   End If
End Sub

Private Sub ChkAllowDiscount_Click()
   If ActiveControl.Name <> ChkAllowDiscount.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkChangeDate_Click()
   If ActiveControl.Name <> ChkChangeDate.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkChangePrice_Click()
   If ActiveControl.Name <> ChkChangePrice.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkCreditSale_Click()
   If ActiveControl.Name <> ChkCreditSale.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkDisableCreditSale_Click()
   If ActiveControl.Name <> ChkDisableCreditSale.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkOpenForm_Click()
   If ActiveControl.Name <> ChkOpenForm.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkPurchaseRePrint_Click()
If ActiveControl.Name <> ChkPurchaseRePrint.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkSaleRePrint_Click()
   If ActiveControl.Name <> ChkSaleRePrint.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkShowPurchasePriceInInvoice_Click()
   If ActiveControl.Name <> ChkShowPurchasePriceInInvoice.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkShowStock_Click()
   If ActiveControl.Name <> ChkShowStock.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkShowSumInSearchSaleInvoice_Click()
   If ActiveControl.Name <> ChkShowSumInSearchSaleInvoice.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub ChkSalePriceMustBeLessThanPurchase_Click()
   If ActiveControl.Name <> ChkSalePriceMustBeLessThanPurchase.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub CmbUsers_Click()
   If CmbUsers.Visible = False Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
'   GetUserSettings
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hwnd, "User Default Settings"
   With CN.Execute("Select * FROM Users order by UserName")
      Do Until .EOF
         CmbUsers.AddItem !UserName
         CmbUsers.ItemData(CmbUsers.NewIndex) = !UserNo
         .MoveNext
      Loop
   End With
   FormStatus = NewMode
   'BtnSave.Visible = Not vReadOnly
   'BtnDelete.Visible = Not vReadOnly
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Dim lngReturnValue As Long
'   If Button = 1 Then
'      Call ReleaseCapture
'      lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
'   End If
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtStoreID.Name: If FunSelectStore(ssFunctionKey, False) = True Then If TxtOrganizationID.Enabled Then TxtOrganizationID.SetFocus Else TxtStoreID.SetFocus
         Case TxtOrganizationID.Name: If FunSelectOrganization(ssFunctionKey, False) = True Then If BtnSave.Enabled Then BtnSave.SetFocus Else TxtOrganizationID.SetFocus
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
'   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
'      MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
'      Exit Sub
'   End If
   If FunValidation = False Then Exit Sub
   CN.BeginTrans
   Set Rs = New ADODB.Recordset
   sSql = "Select * FROM UserRegistry where UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)
   Rs.Open sSql, CN, adOpenStatic, adLockOptimistic
   If Rs.RecordCount = 0 Then
      Rs.AddNew
      Rs!UserNo = CmbUsers.ItemData(CmbUsers.ListIndex)
   End If
   Rs!StoreID = IIf(TxtStoreID.Text = "", Null, TxtStoreID.Text)
   Rs!OrganizationID = IIf(TxtOrganizationID.Text = "", Null, TxtOrganizationID.Text)
   Rs!NoofPrints = Val(TxtNoofPrints.Text)
   Rs!NoofPurPrints = Val(TxtNoofPurPrints.Text)
   Rs!ChangePrice = Abs(ChkChangePrice.Value)
   Rs!CreditSale = Abs(ChkCreditSale.Value)
   Rs!DisableCreditSale = Abs(ChkDisableCreditSale.Value)
   Rs!CustomerPoleDisplay = Abs(ChkShowCustomerPoleDisplay.Value)
   Rs!ShowPurchasePriceInInvoice = Abs(ChkShowPurchasePriceInInvoice.Value)
   Rs!ShowStock = Abs(ChkShowStock.Value)
   Rs!AllowDiscount = Abs(ChkAllowDiscount.Value)
   Rs!ShowSumInSearchSaleInvoice = Abs(ChkShowSumInSearchSaleInvoice.Value)
   Rs!SalePriceMustBeLessThanPurchase = Abs(ChkSalePriceMustBeLessThanPurchase.Value)
   Rs!NotEditingAfterPrinting = Abs(ChkNotEditingAfterPrinting.Value)
   Rs!ChangeDate = Abs(ChkChangeDate.Value)
   Rs!OpenForm = Abs(ChkOpenForm.Value)
   Rs!SaleRePrint = Abs(ChkSaleRePrint.Value)
   Rs!PurchaseRePrint = Abs(ChkPurchaseRePrint.Value)
   Rs!CommPort = IIf(TxtCommPort.Text = "", Null, Val(TxtCommPort.Text))
   Rs!AllowMaximmDiscPer = IIf(TxtAllowMaximmDiscPer.Text = "", 0, Val(TxtAllowMaximmDiscPer.Text))
   Rs!ChangePriceFormOpenAsLogin = Abs(ChkChangePriceFormOpenAsLogin.Value)
   Rs.Update
'   If Trim(TxtStoreID.Text) = "" Then
'      CN.Execute "DELETE FROM UserRegistry where UserNo = " & CmbUsers.ItemData(CmbUsers.ListIndex)
'   End If
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
   On Error GoTo ErrorHandler
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
      If CmbUsers.ListCount > 0 Then CmbUsers.ListIndex = 0
      CmbUsers.Enabled = True
      If CmbUsers.Visible And CmbUsers.Enabled Then CmbUsers.SetFocus
      vIsNewRecord = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
   Case Is = OpenMode
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      CmbUsers.Enabled = False
      If TxtStoreID.Visible And TxtStoreID.Enabled Then TxtStoreID.SetFocus
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
   Set FrmUserDefaultSettings = Nothing
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
   ChkChangePrice.Value = False
   ChkShowCustomerPoleDisplay.Value = False
   ChkCreditSale.Value = False
   ChkDisableCreditSale.Value = False
   ChkShowPurchasePriceInInvoice.Value = False
   ChkShowStock.Value = False
   ChkAllowDiscount.Value = False
   ChkShowSumInSearchSaleInvoice.Value = False
   ChkSalePriceMustBeLessThanPurchase.Value = False
   ChkNotEditingAfterPrinting.Value = False
   ChkChangePriceFormOpenAsLogin.Value = False
   ChkChangeDate.Value = False
   ChkOpenForm.Value = False
   ChkSaleRePrint.Value = False
   ChkPurchaseRePrint.Value = False
   TxtAllowMaximmDiscPer.Text = 0
   TxtNoofPrints.Text = 0
   TxtNoofPurPrints.Text = 0
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
    With CN.Execute(vStrSQL)
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

Private Function FunSelectOrganization(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchOrganization.Show vbModal, Me
        If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganization = False: Exit Function
        TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Organizations where OrganizationID=" & Val(TxtOrganizationID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtOrganizationName.Text = !OrganizationName
          FunSelectOrganization = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectOrganization = False
          .Close
          TxtOrganizationID.Text = ""
          TxtOrganizationName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtOrganizationID_Change()
   If TxtOrganizationID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   If TxtOrganizationName.Text <> "" Then TxtOrganizationName.Text = ""
End Sub

Private Sub TxtOrganizationID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtOrganizationID.Text) = "" Then Exit Sub
   If TxtOrganizationName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectOrganization(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectOrganization(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOrganization_Click()
   If FunSelectOrganization(ssButton, False) = True Then
      If BtnSave.Enabled Then BtnSave.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub
