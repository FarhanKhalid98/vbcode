VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form SchActivityLog 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbUser 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1575
      Width           =   2790
   End
   Begin VB.ComboBox CmbFilter 
      Height          =   315
      Left            =   3195
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1575
      Width           =   3150
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5985
      Left            =   390
      TabIndex        =   0
      Top             =   1950
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
      stylesets(0).Picture=   "SchActivityLog.frx":0000
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
      Left            =   5363
      TabIndex        =   6
      Top             =   8145
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
      MICON           =   "SchActivityLog.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   3195
      TabIndex        =   5
      Top             =   8190
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
      MICON           =   "SchActivityLog.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFromDate 
      Height          =   330
      Left            =   6375
      TabIndex        =   3
      Top             =   1575
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
      Left            =   7815
      TabIndex        =   4
      Top             =   1575
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
      Left            =   360
      TabIndex        =   10
      Top             =   1350
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
      Left            =   3195
      TabIndex        =   9
      Top             =   1350
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
      Left            =   6465
      TabIndex        =   8
      Top             =   1350
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
      Left            =   1920
      TabIndex        =   7
      Top             =   180
      Width           =   1605
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "SchActivityLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vStrSQL As String
Public ParaOutUserNo As String
Dim vOrder As String, vDirection As String, vCol As Byte, vFilter As String, vUser As String

Private Sub BtnClose_Click()
  Me.ParaOutUserNo = ""
  Unload Me
End Sub

Private Sub CmbFilter_Click()
   If CmbFilter.ListIndex = 0 Then
      vFilter = ""
   Else
      vFilter = " and FormType like '%" & CmbFilter.Text & "%'"
   End If
   LoadGrid
End Sub

Private Sub CmbUser_Click()
   If CmbUser.ListIndex = 0 Then
      vUser = ""
   Else
      vUser = " and u.UserNo =" & CmbUser.ItemData(CmbUser.ListIndex)
   End If
   LoadGrid
End Sub

Private Sub DtpFromDate_Change()
   LoadGrid
End Sub

Private Sub DtpToDate_Change()
   LoadGrid
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
   vStrSQL = " Select A.*, UserName from Activitylog A " & _
             " Inner Join Users U on U.UserNO= A.UserNo " & " WHERE EntryDate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'" & vFilter & vUser & vOrder & vDirection
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
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
   vOrder = " order by " & Grid.Columns(ColIndex).DataField
   If vCol = ColIndex Then
      vDirection = IIf(vDirection = " Asc", " Desc", " Asc")
   Else
      vDirection = " Asc"
   End If
   vCol = ColIndex
   LoadGrid
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

