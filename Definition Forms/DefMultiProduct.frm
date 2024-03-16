VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form DefMultiProduct 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   6015
      TabIndex        =   1
      Top             =   8250
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
      MICON           =   "DefMultiProduct.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   4710
      TabIndex        =   0
      Top             =   8250
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
      MICON           =   "DefMultiProduct.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3705
      Left            =   2535
      TabIndex        =   3
      Tag             =   "C"
      ToolTipText     =   "Product Entry"
      Top             =   2430
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   6535
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3705
      Left            =   6360
      TabIndex        =   4
      Tag             =   "C"
      ToolTipText     =   "Product Entry"
      Top             =   2475
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   6535
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin JeweledBut.JeweledButton BtnAddColour 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   2430
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      TX              =   "+"
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
      MICON           =   "DefMultiProduct.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnAddSize 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9105
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "nc"
      ToolTipText     =   "Add New"
      Top             =   2475
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      TX              =   "+"
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
      MICON           =   "DefMultiProduct.frx":0054
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multi Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   2
      Top             =   270
      Width           =   2310
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "DefMultiProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim Rs As ADODB.Recordset
'Dim VStrSQL As String
'Public ParaOutBrandID As String
'Dim vOrder As String, vDirection As String, vCol As Byte
'
'Private Sub BtnClose_Click()
'  Unload Me
'End Sub
'
'Private Sub BtnSelect_Click()
'  On Error GoTo ErrorHandler
''  If Grid.Rows = 0 Then Exit Sub
''  Me.ParaOutBrandID = Rs!BrandID
''  Unload Me
'  Exit Sub
'ErrorHandler:
'  Call ShowErrorMessage
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyEscape Then Call BtnClose_Click
''   If KeyCode = vbKeyReturn Then
''      Select Case ActiveControl.Name
''      Case Grid.Name
''         Call BtnSelect_Click
''      End Select
''   End If
'End Sub
'
'Private Sub Form_Load()
'    On Error GoTo ErrorHandler
'    ShowPicture Me, 2
'    AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
'    SetWindowText Me.hWnd, "Multi Product"
''    Set Rs = New ADODB.Recordset
''    VStrSQL = "SELECT * FROM Parties WHERE PartyType='V'"
''    Rs.Open VStrSQL, CN, adOpenStatic, adLockReadOnly
''    Set Grid.DataSource = Rs
''    Grid.Columns("ID").DataField = "BrandID"
''    Grid.Columns("Name").DataField = "BrandName"
''    Grid.Columns("Address").DataField = "Address"
''    Grid.Columns("City").DataField = "City"
'    LoadData
'    Exit Sub
'ErrorHandler:
'  Call ShowErrorMessage
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   On Error GoTo ErrorHandler
'   Dim frmObj As Object
'   For Each frmObj In Forms
'       Set frmObj = Nothing
'   Next
'   Set Rs = Nothing
'Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub
'
'Private Sub ImgExit_Click()
'   Unload Me
'End Sub
'
'Private Sub LoadData()
'   On Error GoTo ErrorHandler
''   VStrSQL = "SELECT * FROM Brands where 1=1 " & IIf(Trim(TxtBrandName.Text) = "", "", " and BrandName like '%" & TxtBrandName.Text & "%'") & vOrder & vDirection
''   If Rs.State = adStateOpen Then Rs.Close
''   Rs.Open VStrSQL, CN, adOpenStatic, adLockReadOnly
''   Set Grid.DataSource = Rs
''   Grid.Columns("ID").DataField = "BrandID"
''   Grid.Columns("Name").DataField = "BrandName"
'
'   With CN.Execute("Select * FROM Colors")
'      Do Until .EOF
'         Set lvwItem = ListView1.ListItems.Add(, , .Fields.Item("Color").Value)
'         .MoveNext
'      Loop
'   End With
'
'   With CN.Execute("Select * FROM Sizes")
'      Do Until .EOF
'         Set lvwItem = ListView1.ListItems.Add(, , .Fields.Item("Size").Value)
'         .MoveNext
'      Loop
'   End With
'
'
''   Set Item = ListView1.ListItems.Add(, , TxtItemNo.Text)
'
'   Exit Sub
'ErrorHandler:
'   Call ShowErrorMessage
'End Sub
'
