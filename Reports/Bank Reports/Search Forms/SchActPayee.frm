VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form SchActPayee 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbfilter 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8010
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1590
      Width           =   2355
   End
   Begin VB.TextBox TxtFilter 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4035
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1560
      Width           =   3915
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6015
      Left            =   4035
      TabIndex        =   0
      Top             =   1875
      Width           =   3930
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
      stylesets(0).ForeColor=   -2147483634
      stylesets(0).BackColor=   -2147483635
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
      stylesets(0).Picture=   "SchActPayee.frx":0000
      AllowAddNew     =   -1  'True
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
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "Account Payee ID"
      Columns(0).Name =   "ID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Account Payee Name"
      Columns(1).Name =   "Name"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   6932
      _ExtentY        =   10610
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
   Begin JeweledBut.JeweledButton CmdSelect 
      Default         =   -1  'True
      Height          =   420
      Left            =   4710
      TabIndex        =   4
      Top             =   8055
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
      MICON           =   "SchActPayee.frx":001C
      BC              =   16777215
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdClose 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   6015
      TabIndex        =   5
      Top             =   8055
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Cancel"
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
      MICON           =   "SchActPayee.frx":0038
      BC              =   16777215
      FC              =   0
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   180
      Width           =   1005
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
      Caption         =   "Account Payee ID"
      Height          =   195
      Left            =   8040
      TabIndex        =   3
      Top             =   1365
      Width           =   1305
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Payee Name"
      Height          =   195
      Left            =   4035
      TabIndex        =   2
      Top             =   1365
      Width           =   1560
   End
End
Attribute VB_Name = "SchActPayee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rs As New ADODB.Recordset
Public ParaOutID As String
Public ParaOutName As String
Public ParaInWhereClause As String
Dim vOrder As String, vDirection As String, vCol As Byte

Private Sub cmbfilter_click()
   Dim vwhere As String
   On Error GoTo ErrorHandler
   If cmbfilter.ListIndex > 0 Then
      vwhere = "where " & Rs.Fields.Item(0).Name & "= '" & cmbfilter.Text & " ' "
   End If
   Set Rs = CN.Execute("select PartyID, PartyName  from Parties " & vwhere)
   Set Grid.DataSource = Rs
   Grid.Columns("ID").DataField = Rs.Fields.Item(0).Name
   Grid.Columns("Name").DataField = Rs.Fields.Item(1).Name
   
'   Grid.Columns("Name").DataField = "Job_DESC"
   
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

 Private Sub CmdClose_Click()
   Me.ParaOutID = ""
   Me.ParaOutName = ""
     Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      CmdSelect_Click
   End If
End Sub

Private Sub Grid_DblClick()
  If Grid.Rows > 0 Then CmdSelect_Click
End Sub

Private Sub CmdSelect_Click()
   On Error GoTo ErrorHandler
   If Grid.Rows = 0 Then Exit Sub
   Me.ParaOutID = Rs.Fields(0).Value
   Me.ParaOutName = Rs.Fields(1).Value
   Unload Me
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Search"
   cmbfilter.AddItem "All Products", 0
   With CN.Execute("select PartyID from Parties")
      Do Until .EOF
         cmbfilter.AddItem .Fields(0).Value
         .MoveNext
      Loop
   End With
   If cmbfilter.ListCount > 0 Then cmbfilter.ListIndex = 0
   Me.ParaOutID = ""
   Me.ParaOutName = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If Rs.RecordCount > 0 Then
      'TxtFilter.Text = Grid.Columns("Name").Text
      cmbfilter.Text = Grid.Columns("ID").Text
   End If
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtFilter_Change()
  On Error GoTo ErrorHandler
  If Trim(TxtFilter.Text) = "" Then Grid.MoveFirst: Exit Sub
  Rs.Find Rs.Fields.Item(1).Name & " like '" & Replace(TxtFilter.Text, "'", "''") & "%'", , adSearchForward, 1
  If Rs.EOF Then Grid.MoveLast
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

