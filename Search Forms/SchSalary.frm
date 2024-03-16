VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form SchSalary 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6330
      Left            =   2067
      TabIndex        =   4
      Top             =   3195
      Width           =   10635
      ScrollBars      =   2
      _Version        =   196616
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RecordSelectors =   0   'False
      stylesets.count =   2
      stylesets(0).Name=   "ColUrdu"
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "SchSalary.frx":0000
      stylesets(1).Name=   "ColEng"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "SchSalary.frx":001C
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
      RowHeight       =   609
      Columns.Count   =   7
      Columns(0).Width=   1799
      Columns(0).Caption=   "Emp ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).StyleSet=   "ColEng"
      Columns(1).Width=   3863
      Columns(1).Caption=   "Emp Name"
      Columns(1).Name =   "Name"
      Columns(1).Alignment=   2
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).StyleSet=   "ColEng"
      Columns(2).Width=   3863
      Columns(2).Caption=   "Father Name"
      Columns(2).Name =   "FName"
      Columns(2).Alignment=   2
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 4"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).StyleSet=   "ColEng"
      Columns(3).Width=   3149
      Columns(3).Caption=   "Designation"
      Columns(3).Name =   "Designation"
      Columns(3).Alignment=   2
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 2"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).StyleSet=   "ColEng"
      Columns(4).Width=   1429
      Columns(4).Caption=   "Month"
      Columns(4).Name =   "Date"
      Columns(4).Alignment=   2
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 5"
      Columns(4).DataType=   7
      Columns(4).NumberFormat=   "MMM-yyyy"
      Columns(4).FieldLen=   256
      Columns(4).StyleSet=   "ColEng"
      Columns(5).Width=   2302
      Columns(5).Caption=   "Entry Date"
      Columns(5).Name =   "EntryDate"
      Columns(5).Alignment=   2
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 6"
      Columns(5).DataType=   8
      Columns(5).NumberFormat=   "dd/MM/yyyy"
      Columns(5).FieldLen=   256
      Columns(5).StyleSet=   "ColEng"
      Columns(6).Width=   1746
      Columns(6).Caption=   "Amount"
      Columns(6).Name =   "Total"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 3"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).StyleSet=   "ColEng"
      TabNavigation   =   1
      _ExtentX        =   18759
      _ExtentY        =   11165
      _StockProps     =   79
      BackColor       =   15724527
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   7399
      TabIndex        =   7
      Top             =   9810
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
      MICON           =   "SchSalary.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   6094
      TabIndex        =   6
      Top             =   9810
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
      MICON           =   "SchSalary.frx":0054
      BC              =   14737632
      FC              =   0
   End
   Begin MSComCtl2.DTPicker DtpFrom 
      Height          =   345
      Left            =   7557
      TabIndex        =   3
      Top             =   2880
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "MMM-yyyy"
      Format          =   123076611
      CurrentDate     =   38718
   End
   Begin MSComCtl2.DTPicker DtpTo 
      Height          =   345
      Left            =   8847
      TabIndex        =   5
      Top             =   2880
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "MMM-yyyy"
      Format          =   123076611
      CurrentDate     =   38718
   End
   Begin SITextBox.Txt TxtID 
      Height          =   315
      Left            =   2082
      TabIndex        =   0
      Top             =   2880
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
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
   End
   Begin SITextBox.Txt TxtName 
      Height          =   315
      Left            =   3102
      TabIndex        =   1
      Top             =   2880
      Width           =   2200
      _ExtentX        =   3889
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
      Masked          =   5
   End
   Begin SITextBox.Txt TxtFName 
      Height          =   315
      Left            =   5307
      TabIndex        =   2
      Top             =   2880
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   30
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
      TabIndex        =   12
      Top             =   270
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "--------------------Month--------------------"
      Height          =   195
      Left            =   7557
      TabIndex        =   11
      Top             =   2670
      Width           =   2250
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
      Height          =   195
      Left            =   2082
      TabIndex        =   10
      Top             =   2670
      Width           =   525
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
      Height          =   195
      Left            =   3102
      TabIndex        =   9
      Top             =   2670
      Width           =   780
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      Height          =   195
      Left            =   5307
      TabIndex        =   8
      Top             =   2700
      Width           =   915
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   12994
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image ImgMin 
      Height          =   315
      Left            =   12499
      Top             =   1710
      Width           =   375
   End
End
Attribute VB_Name = "SchSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vStrSQL As String
Public ParaOutEmpID As String
Public ParaOutDate As String

Private Sub BtnClose_Click()
  Me.ParaOutEmpID = ""
  Me.ParaOutDate = ""
  Unload Me
End Sub

Private Sub BtnSelect_Click()
  On Error GoTo ErrorHandler
  If Grid.Rows = 0 Then Exit Sub
  Me.ParaOutEmpID = Rs!EmpID
  Me.ParaOutDate = Rs!SalaryMonth
  Unload Me
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub DtpFrom_Change()
   LoadGrid
End Sub

Private Sub DtpTo_Change()
   LoadGrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyEscape
      Call BtnClose_Click
      KeyCode = 0
   Case vbKeyReturn
         Call BtnSelect_Click
         KeyCode = 0
   Case vbKeyDown
      If UCase(ActiveControl.Name) = UCase("txt*") Or UCase(ActiveControl.Name) = UCase("dtp*") Then
         Grid.SetFocus
         KeyCode = 0
      End If
   End Select
End Sub

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   vStrSQL = " Select s.*, Empname,  SalaryMonth" & vbCrLf _
            & " , case when previous < 0 then ttlsalary-less-previous else ttlsalary-less end as total" & vbCrLf _
            & " FROM salaries s inner join Employees e on s.EmpID = e.EmpID" _
            & " WHERE SalaryMonth >='" & DtpFrom.Value & "' and SalaryMonth < '" & DateAdd("m", 1, DtpTo.Value) & "'" & IIf(Trim(TxtName.Text) = "", "", " and e.empName like '%" & TxtName.Text & "%'") & IIf(Trim(TxtID.Text) = "", "", " and e.EmpID ='" & TxtID.Text & "'") & IIf(Trim(TxtFName.Text) = "", "", " and e.FName like '%" & TxtFName.Text & "%' ")
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Set Grid.DataSource = Rs
   Grid.SelBookmarks.RemoveAll
   Grid.SelBookmarks.Add Grid.Bookmark
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Search"
   
   Grid.Columns("ID").DataField = "EmpID"
   Grid.Columns("Name").DataField = "EmpName"
   'Grid.Columns("FName").DataField = "FName"
   Grid.Columns("Date").DataField = "SalaryMonth"
   Grid.Columns("EntryDate").DataField = "EntryDate"
   Grid.Columns("Designation").DataField = "Designation"
   Grid.Columns("Total").DataField = "Total"
   Me.ParaOutEmpID = ""
   Me.ParaOutDate = ""
   DtpFrom.Value = Date
   DtpTo.Value = Date
   DtpFrom.Value = DtpFrom.Value - 30
   DtpFrom.Day = 1
   DtpTo.Day = 1
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
  If Grid.Rows > 0 Then Call BtnSelect_Click
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case Asc("a") To Asc("z"), Asc("A") To Asc("Z"), vbKey0 To vbKey9
      TxtID.Text = Chr(KeyAscii): TxtID.SetFocus
   End Select
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub ImgMin_Click()
   Me.WindowState = 1
End Sub
Private Sub TxtFName_Change()
On Error GoTo ErrorHandler
'   If Trim(TxtName.Text) = "" Then Grid.MoveFirst: Exit Sub
'   Rs.Find "Name like '%" & TxtName.Text & "%'", , adSearchForward, 1
'   If Rs.EOF Then Grid.MoveLast
'   Grid.SelBookmarks.RemoveAll
'   Grid.SelBookmarks.Add Grid.Bookmark
    Call LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtID_Change()
   On Error GoTo ErrorHandler
'   If Trim(TxtID.Text) = "" Then Grid.MoveFirst: Exit Sub
'   Rs.Find "EmpID like '" & TxtID.Text & "%'", , adSearchForward, 1
'   If Rs.EOF Then Grid.MoveLast
'   Grid.SelBookmarks.RemoveAll
'   Grid.SelBookmarks.Add Grid.Bookmark
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtName_Change()
   On Error GoTo ErrorHandler
'   If Trim(TxtName.Text) = "" Then Grid.MoveFirst: Exit Sub
'   Rs.Find "Name like '%" & TxtName.Text & "%'", , adSearchForward, 1
'   If Rs.EOF Then Grid.MoveLast
'   Grid.SelBookmarks.RemoveAll
'   Grid.SelBookmarks.Add Grid.Bookmark
    Call LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
