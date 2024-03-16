VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmActivityLog 
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
   Begin VB.ComboBox CmbAction 
      Height          =   315
      ItemData        =   "FrmActivityLog.frx":0000
      Left            =   10800
      List            =   "FrmActivityLog.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2385
      Width           =   2250
   End
   Begin VB.ComboBox CmbUser 
      Height          =   315
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2385
      Width           =   2790
   End
   Begin VB.ComboBox CmbFilter 
      Height          =   315
      Left            =   4725
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2385
      Width           =   3150
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5985
      Left            =   1920
      TabIndex        =   0
      Top             =   2760
      Width           =   11550
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "Select"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   13817275
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
      stylesets(0).Picture=   "FrmActivityLog.frx":0004
      AllowUpdate     =   0   'False
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
      RowNavigation   =   1
      ForeColorEven   =   0
      BackColorEven   =   15724527
      BackColorOdd    =   16777215
      RowHeight       =   423
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   8
      Columns(0).Width=   1191
      Columns(0).Caption=   "User ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2434
      Columns(1).Caption=   "User Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3043
      Columns(2).Caption=   "Form Type"
      Columns(2).Name =   "FormType"
      Columns(2).DataField=   "Column 7"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3069
      Columns(3).Caption=   "Entry Date"
      Columns(3).Name =   "EntryDate"
      Columns(3).DataField=   "Column 2"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   7858
      Columns(4).Caption=   "Description"
      Columns(4).Name =   "Description"
      Columns(4).DataField=   "Column 3"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   820
      Columns(5).Caption=   "New"
      Columns(5).Name =   "IsNew"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   4
      Columns(5).FieldLen=   256
      Columns(5).Style=   2
      Columns(6).Width=   714
      Columns(6).Caption=   "Edit"
      Columns(6).Name =   "IsEdit"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   4
      Columns(6).FieldLen=   256
      Columns(6).Style=   2
      Columns(7).Width=   767
      Columns(7).Caption=   "Del"
      Columns(7).Name =   "IsDelete"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   4
      Columns(7).FieldLen=   256
      Columns(7).Style=   2
      TabNavigation   =   1
      _ExtentX        =   20373
      _ExtentY        =   10557
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
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   6893
      TabIndex        =   7
      Top             =   8955
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
      MICON           =   "FrmActivityLog.frx":0020
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   4725
      TabIndex        =   6
      Top             =   9000
      Visible         =   0   'False
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
      MICON           =   "FrmActivityLog.frx":003C
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFromDate 
      Height          =   330
      Left            =   7905
      TabIndex        =   3
      Top             =   2385
      Width           =   1395
      _Version        =   65543
      _ExtentX        =   2461
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpToDate 
      Height          =   330
      Left            =   9345
      TabIndex        =   4
      Top             =   2385
      Width           =   1395
      _Version        =   65543
      _ExtentX        =   2461
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
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
   Begin JeweledBut.JeweledButton BtnRecyclebin 
      Height          =   420
      Left            =   11745
      TabIndex        =   13
      Top             =   1710
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Recycle Bin"
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
      MICON           =   "FrmActivityLog.frx":0058
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnRefresh 
      Height          =   420
      Left            =   10440
      TabIndex        =   14
      Top             =   1710
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Refresh"
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
      MICON           =   "FrmActivityLog.frx":0074
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Actions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10800
      TabIndex        =   12
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1890
      TabIndex        =   11
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Form Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4725
      TabIndex        =   10
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "-------From Date To Date ---------"
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
      Left            =   7995
      TabIndex        =   9
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activity Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2700
      TabIndex        =   8
      Top             =   270
      Width           =   1605
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "FrmActivityLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vStrSQL As String
Public ParaOutUserNo As String
Dim vOrder As String, vDirection As String, vCol As Byte, vFilter As String, vUser As String, vAction As String

Private Sub BtnClose_Click()
   On Error GoTo ErrorHandler
   Me.ParaOutUserNo = ""
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnRecyclebin_Click()
FrmRecycleBin.Show
End Sub

Private Sub BtnRefresh_Click()
   LoadGrid
End Sub

Private Sub CmbAction_Click()
   On Error GoTo ErrorHandler
   Select Case CmbAction.ListIndex
   Case 0
      vAction = ""
   Case 1
      vAction = " and IsNew = 1"
   Case 2
      vAction = " and a.IsEdit = 1"
   Case 3
      vAction = " and a.IsDelete = 1"
   Case 4
      vAction = " and a.IsClear = 1"
   End Select
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbFilter_Click()
   On Error GoTo ErrorHandler
   If CmbFilter.ListIndex = 0 Then
      vFilter = ""
   Else
      vFilter = " and FormType like '%" & CmbFilter.Text & "%'"
   End If
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbUser_Click()
   On Error GoTo ErrorHandler
   If CmbUser.ListIndex = 0 Then
      vUser = ""
   Else
      vUser = " and u.UserNo =" & CmbUser.ItemData(CmbUser.ListIndex)
   End If
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub DtpFromDate_Change()
   On Error GoTo ErrorHandler
   If DtpFromDate.IsDateValid = False Then Exit Sub
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub DtpToDate_Change()
   On Error GoTo ErrorHandler
   If DtpToDate.IsDateValid = False Then Exit Sub
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyEscape Then Call BtnClose_Click
'   If KeyCode = vbKeyReturn Then
'      Select Case ActiveControl.Name
'      Case Grid.Name, TxtUserName.Name
'         Call BtnSelect_Click
'      End Select
'   End If
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hwnd, "Activity Log"
   CmbFilter.Clear
   CmbFilter.AddItem "-- All Types --"
   With CN.Execute("Select Distinct FormType FROM Activitylog")
      Do Until .EOF
          CmbFilter.AddItem !FormType
          .MoveNext
      Loop
   End With
   CmbUser.Clear
   CmbUser.AddItem "-- All Users --"
   With CN.Execute("Select * FROM Users")
      Do Until .EOF
          CmbUser.AddItem !UserName
          CmbUser.ItemData(CmbUser.NewIndex) = !UserNo
          .MoveNext
      Loop
   End With
   
   CmbAction.Clear
   CmbAction.AddItem "-- All Actions --"
   CmbAction.AddItem "New"
   CmbAction.AddItem "Edit"
   CmbAction.AddItem "Delete"
   CmbAction.AddItem "Clear"
   CmbAction.ListIndex = 0
   
   DtpFromDate.DateValue = Date - 10
   DtpToDate.DateValue = Date
   CmbFilter.ListIndex = 0
   CmbUser.ListIndex = 0
   LoadGrid
   Me.ParaOutUserNo = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   vStrSQL = " Select A.*, UserName from Activitylog A " & _
             " Inner Join Users U on U.UserNO= A.UserNo " & " WHERE EntryDate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "' and Description <> '' " & vAction & vFilter & vUser & vOrder & vDirection
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Grid.Columns("ID").DataField = "UserNo"
   Grid.Columns("Name").DataField = "UserName"
   Grid.Columns("FormType").DataField = "FormType"
   Grid.Columns("EntryDate").DataField = "EntryDate"
   Grid.Columns("Description").DataField = "Description"
   Grid.Columns("IsNew").DataField = "IsNew"
   Grid.Columns("IsEdit").DataField = "IsEdit"
   Grid.Columns("IsDelete").DataField = "IsDelete"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo ErrorHandler
   vOrder = " order by " & Grid.Columns(ColIndex).DataField
   If vCol = ColIndex Then
      vDirection = IIf(vDirection = " Asc", " Desc", " Asc")
   Else
      vDirection = " Asc"
   End If
   vCol = ColIndex
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   On Error GoTo ErrorHandler
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

