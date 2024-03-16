VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmActivityLogBin 
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
      ItemData        =   "FrmActivityLogBin.frx":0000
      Left            =   10800
      List            =   "FrmActivityLogBin.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2385
      Width           =   4095
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
      Left            =   165
      TabIndex        =   0
      Top             =   2760
      Width           =   15105
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
      stylesets(0).Picture=   "FrmActivityLogBin.frx":0004
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
      Columns.Count   =   10
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "Activity ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3228
      Columns(1).Caption=   "Entry Date"
      Columns(1).Name =   "EntryDate"
      Columns(1).DataField=   "Column 2"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1270
      Columns(2).Caption=   "User No"
      Columns(2).Name =   "UserID"
      Columns(2).DataField=   "Column 5"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2328
      Columns(3).Caption=   "User Name"
      Columns(3).Name =   "Name"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 1"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "FormNo"
      Columns(4).Name =   "FormNo"
      Columns(4).DataField=   "Column 9"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3678
      Columns(5).Caption=   "Form Type"
      Columns(5).Name =   "FormType"
      Columns(5).DataField=   "Column 7"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1640
      Columns(6).Caption=   "TXN ID"
      Columns(6).Name =   "TransactionID"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1746
      Columns(7).Caption=   "TXN Date"
      Columns(7).Name =   "TransactionDate"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   7
      Columns(7).NumberFormat=   "dd/mm/yy"
      Columns(7).FieldLen=   256
      Columns(8).Width=   6773
      Columns(8).Caption=   "Description"
      Columns(8).Name =   "Description"
      Columns(8).DataField=   "Column 3"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   5450
      Columns(9).Caption=   "Action Name"
      Columns(9).Name =   "ActionName"
      Columns(9).DataField=   "Column 8"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   26644
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
      MICON           =   "FrmActivityLogBin.frx":0020
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
      MICON           =   "FrmActivityLogBin.frx":003C
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
      MICON           =   "FrmActivityLogBin.frx":0058
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
      MICON           =   "FrmActivityLogBin.frx":0074
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
Attribute VB_Name = "FrmActivityLogBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vStrSQL, vActionNo As String
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
FrmRecycleBinNew.Show
End Sub

Private Sub BtnRefresh_Click()
   LoadGrid
End Sub

Private Sub CmbAction_Click()
   On Error GoTo ErrorHandler
   If CmbAction.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbAction.Name Then Exit Sub
   If CmbAction.ListIndex = 0 Then
      vAction = ""
   Else
      vAction = " and AB.ActionNo =" & CmbAction.ItemData(CmbAction.ListIndex)
   End If
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbFilter_Click()
   On Error GoTo ErrorHandler
   If CmbFilter.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbFilter.Name Then Exit Sub
   If CmbFilter.ListIndex = 0 Then
      vFilter = ""
   Else
      vFilter = " and ALB.FormNo = " & CmbFilter.ItemData(CmbFilter.ListIndex)
   End If
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbUser_Click()
   On Error GoTo ErrorHandler
   If CmbUser.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbUser.Name Then Exit Sub
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
   With CN.Execute("Select * FROM " & vBinDataBase & ".dbo.FormBin Order by FormNo Asc")
      Do Until .EOF
          CmbFilter.AddItem !FormType
          CmbFilter.ItemData(CmbFilter.NewIndex) = !FormNo
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
   
   vActionNo = ObjRegistry.ByDefaultActionNo
   CmbAction.Clear
   CmbAction.AddItem "-- All Actions --"
   With CN.Execute("Select * FROM " & vBinDataBase & ".dbo.ActionBin  " & IIf(vActionNo = "", "", " Where ActionNo in (" & vActionNo & " )"))
      Do Until .EOF
          CmbAction.AddItem !ActionName
          CmbAction.ItemData(CmbAction.NewIndex) = !ActionNo
          .MoveNext
      Loop
   End With
   
   DtpFromDate.DateValue = Date - 10
   DtpToDate.DateValue = Date
   CmbFilter.ListIndex = 0
   CmbUser.ListIndex = 0
   CmbAction.ListIndex = 0
   vOrder = " Order by ActivityID Desc"
   LoadGrid
   Me.ParaOutUserNo = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   vStrSQL = " Select ALB.*, CONVERT (varchar, ActivityDate,103 ) + ' ' +  CONVERT (varchar, ActivityDate,108 )ActivityDate2, ActionName, FormType, UserName from " & vBinDataBase & ".dbo.ActivitylogBin ALB " & vbCrLf _
             & " Inner Join " & vBinDataBase & ".dbo.ActionBin AB ON AB.ActionNo = ALB.ActionNo " & vbCrLf _
             & " Inner Join " & vBinDataBase & ".dbo.FormBin FB ON FB.FormNo = ALB.FormNo " & vbCrLf _
             & " Inner Join Users U on U.UserNO = ALB.UserNo " & vbCrLf _
             & IIf(vActionNo = "", "Where 1=1 ", " Where AB.ActionNo in (" & vActionNo & " )") & vbCrLf _
             & " And ActivityDate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "' and TransactionInfo <> '' " & vAction & vFilter & vUser & vOrder & vDirection
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Grid.Columns("UserID").DataField = "UserNo"
   Grid.Columns("Name").DataField = "UserName"
   Grid.Columns("FormNo").DataField = "FormNo"
   Grid.Columns("FormType").DataField = "FormType"
   Grid.Columns("ID").DataField = "ActivityID"
   Grid.Columns("EntryDate").DataField = "ActivityDate2"
   Grid.Columns("TransactionID").DataField = "TransactionID"
   Grid.Columns("TransactionDate").DataField = "TransactionDate"
   Grid.Columns("Description").DataField = "TransactionInfo"
   Grid.Columns("ActionName").DataField = "ActionName"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo ErrorHandler
   Select Case Grid.Columns(ColIndex).DataField
      Case "UserName"
            vOrder = " order by U.UserName "
      Case "FormType"
            vOrder = " order by FormType "
      Case "ActionName"
            vOrder = " order by ActionName "
      Case Else
            vOrder = " order by ALB." & Grid.Columns(ColIndex).DataField
   End Select
   
   
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

