VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form SchEmpLeave 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F8E8D6&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   Icon            =   "SchEmpLeave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtEmpID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3293
      TabIndex        =   6
      Top             =   3128
      Width           =   1095
   End
   Begin VB.TextBox TxtEmployeeName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4388
      TabIndex        =   4
      Top             =   3128
      Width           =   3000
   End
   Begin MSComCtl2.DTPicker DtpFrom 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   7388
      TabIndex        =   1
      Top             =   3128
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "MMM-yyyy"
      Format          =   122814467
      CurrentDate     =   38244
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5595
      Left            =   3293
      TabIndex        =   0
      Top             =   3488
      Width           =   6480
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
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
      stylesets(0).Picture=   "SchEmpLeave.frx":0ECA
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
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
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   3
      Columns(0).Width=   1905
      Columns(0).Caption=   "Employee ID"
      Columns(0).Name =   "EmpID"
      Columns(0).DataField=   "Column 2"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   6376
      Columns(1).Caption=   "Employee Name"
      Columns(1).Name =   "EmployeeName"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 4"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2514
      Columns(2).Caption=   "Leave Date"
      Columns(2).Name =   "ApplicationDate"
      Columns(2).Alignment=   2
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 1"
      Columns(2).DataType=   8
      Columns(2).NumberFormat=   "dd/MM/yyyy"
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   11430
      _ExtentY        =   9869
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
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   4883
      TabIndex        =   2
      Top             =   9788
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
      MICON           =   "SchEmpLeave.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   6203
      TabIndex        =   3
      Top             =   9788
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
      MICON           =   "SchEmpLeave.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DTpTo 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8528
      TabIndex        =   9
      Top             =   3128
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "MMM-yyyy"
      Format          =   122814467
      CurrentDate     =   38244
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      Left            =   3000
      TabIndex        =   10
      Top             =   270
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "-------------Leave Date-------------"
      Height          =   195
      Left            =   7493
      TabIndex        =   8
      Top             =   2918
      Width           =   2010
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   3293
      TabIndex        =   7
      Top             =   2918
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   11798
      Top             =   1703
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      Height          =   195
      Left            =   4628
      TabIndex        =   5
      Top             =   2918
      Width           =   1155
   End
End
Attribute VB_Name = "SchEmpLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vOrder As String, vDirection As String, vCol As Byte
Public ParaOutEmpID As Long
Public ParaOutApplicationDate As String

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   Set Rs = New ADODB.Recordset
   Dim vSQL As String
   vSQL = "SELECT  EmpName as EmployeeName" & vbCrLf _
         + " , EL.EmpID, FromDate" & vbCrLf _
         + " FROM EmpLeaves EL INNER JOIN" & vbCrLf _
         + " Employees Emp on Emp.EmpID = EL.EmpID " & vbCrLf _
         + " WHERE FromDate > '" & DtpFrom.Value & "' and FromDate < '" & DateAdd("m", 1, DtpTo.Value) & "'" & vOrder & vDirection
   Rs.Open vSQL, CN, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Grid.Columns("ApplicationDate").DataField = "FromDate"
   Grid.Columns("EmpID").DataField = "EmpID"
   Grid.Columns("EmployeeName").DataField = "EmployeeName"
 Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
  Me.ParaOutEmpID = 0
  Me.ParaOutApplicationDate = Date
  Unload Me
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
  If Grid.Rows = 0 Then Exit Sub
  Me.ParaOutEmpID = Rs!EmpID
  Me.ParaOutApplicationDate = Rs!FromDate
  Unload Me
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub DtpFrom_Change()
   Call LoadGrid
End Sub

Private Sub DtpTo_Change()
   Call LoadGrid
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case ActiveControl.Name
   Case TxtEmpID.Name
      Call NonNumeric(KeyAscii, ActiveControl, True)
   End Select
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Search"
   DtpFrom.Value = Date
   Me.ParaOutEmpID = 0
   Me.ParaOutApplicationDate = Date
   DtpFrom.Value = Date
   DtpTo.Value = Date
   DtpFrom.Value = DtpFrom.Value - 30
   Call LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      Select Case ActiveControl.Name
      Case Grid.Name, TxtEmpID.Name, DtpFrom.Name
         Call BtnSelect_Click
      End Select
   End If
End Sub

Private Sub Grid_DblClick()
  If Grid.Rows > 0 Then BtnSelect_Click
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

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9
      TxtEmpID.Text = Chr(KeyAscii): TxtEmpID.SelStart = Len(TxtEmpID.Text): TxtEmpID.SetFocus
   Case Asc("a") To Asc("z"), Asc("A") To Asc("Z")
      TxtEmployeeName.Text = Chr(KeyAscii): TxtEmployeeName.SelStart = Len(TxtEmployeeName.Text): TxtEmployeeName.SetFocus
   End Select
End Sub

Private Sub Image1_Click()
   Unload Me
End Sub

Private Sub TxtEmployeeName_Change()
On Error GoTo ErrorHandler
   If Trim(TxtEmployeeName.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "EmployeeName like '" & TxtEmployeeName.Text & "%'", , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtEmpID_Change()
 On Error GoTo ErrorHandler
   If Trim(TxtEmpID.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "EmpID = " & TxtEmpID.Text, , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

