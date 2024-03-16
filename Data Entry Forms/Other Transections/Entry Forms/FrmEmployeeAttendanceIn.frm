VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmEmployeeAttendanceIn 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmEmployeeAttendanceIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtFine 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   10530
      MaxLength       =   4
      TabIndex        =   22
      Top             =   6578
      Width           =   1350
   End
   Begin VB.TextBox TxtFinePerMinute 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   9135
      MaxLength       =   4
      TabIndex        =   20
      Top             =   6578
      Width           =   1350
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9278
      TabIndex        =   9
      Top             =   7921
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
      MICON           =   "FrmEmployeeAttendanceIn.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7928
      TabIndex        =   5
      Top             =   7928
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
      MICON           =   "FrmEmployeeAttendanceIn.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5318
      TabIndex        =   7
      Top             =   7921
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
      MICON           =   "FrmEmployeeAttendanceIn.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10598
      TabIndex        =   10
      Top             =   7921
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
      MICON           =   "FrmEmployeeAttendanceIn.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6608
      TabIndex        =   6
      Top             =   7928
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
      MICON           =   "FrmEmployeeAttendanceIn.frx":0F3A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   3998
      TabIndex        =   8
      Top             =   7921
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Print"
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
      MICON           =   "FrmEmployeeAttendanceIn.frx":0F56
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmpID 
      Height          =   315
      Left            =   5393
      TabIndex        =   0
      Top             =   4531
      Width           =   1320
      _ExtentX        =   2328
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
   End
   Begin SITextBox.Txt TxtEmpName 
      Height          =   315
      Left            =   7073
      TabIndex        =   2
      Top             =   4531
      Width           =   4350
      _ExtentX        =   7673
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
   Begin JeweledBut.JeweledButton BtnEmployee 
      Height          =   330
      Left            =   6713
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4516
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
      MICON           =   "FrmEmployeeAttendanceIn.frx":0F72
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtAttendID 
      Height          =   315
      Left            =   3488
      TabIndex        =   14
      Top             =   3023
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
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
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpAttendDate 
      Height          =   315
      Left            =   7692
      TabIndex        =   3
      Top             =   5498
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
   Begin MSComCtl2.DTPicker DtpTimeIn 
      Height          =   315
      Left            =   7703
      TabIndex        =   4
      Top             =   6578
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Format          =   102170626
      UpDown          =   -1  'True
      CurrentDate     =   39805.5416666667
   End
   Begin MSComCtl2.DTPicker DtpOfficeTimeIn 
      Height          =   315
      Left            =   6360
      TabIndex        =   18
      Top             =   6600
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   90505218
      UpDown          =   -1  'True
      CurrentDate     =   39805.4166666667
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Fine"
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
      Left            =   10530
      TabIndex        =   23
      Top             =   6360
      Width           =   435
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fine Per Minute"
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
      Left            =   9135
      TabIndex        =   21
      Top             =   6360
      Width           =   1350
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Office Time"
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
      Left            =   6345
      TabIndex        =   19
      Top             =   6390
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date In"
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
      Left            =   7692
      TabIndex        =   17
      Top             =   5273
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attend. Time In"
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
      Left            =   7695
      TabIndex        =   16
      Top             =   6360
      Width           =   1320
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Attend ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3488
      TabIndex        =   15
      Top             =   2783
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7103
      TabIndex        =   13
      Top             =   4306
      Width           =   1320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5393
      TabIndex        =   12
      Top             =   4306
      Width           =   1005
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Attendance In"
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
      TabIndex        =   11
      Top             =   270
      Width           =   4245
   End
   Begin VB.Image ImgExit 
      Height          =   345
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu MniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmEmployeeAttendanceIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
Dim vCounter As Integer
Dim RsReport As New ADODB.Recordset
Dim Flag As Boolean
Dim vMinute As Integer
Dim sSql As String
Dim vStrSQL As String
'----------------------------------
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
   If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsDelete = False Then
      MsgBox "You are not authorized to delete a posted record", vbCritical, "Error"
      Exit Sub
   End If
   If MsgBox("Do you want to remove this record?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
   cn.BeginTrans
    cn.Execute "Delete from EmpAttendance where EmpID = " & Val(TxtEmpID.Text) & " And AttendDate = '" & DtpAttendDate.DateValue & "'"
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If cn.Errors.Count > 0 Then cn.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Sub BtnEmployee_Click()
   If FunSelectEmployee(ssButton, False) = True Then
      If DtpAttendDate.Enabled = True Then DtpAttendDate.SetFocus
   Else
      TxtEmpID.SetFocus
   End If
End Sub

Private Sub BtnOpen_Click()
 SchEmpAttendIn.Show vbModal
   If SchEmpAttendIn.ParaOutEmpID <> 0 Then
      TxtEmpID.Text = SchEmpAttendIn.ParaOutEmpID
      DtpAttendDate.DateValue = SchEmpAttendIn.ParaOutAttendDate
      GetEmployeeAttendace
   End If '
End Sub

Private Sub BtnPrint_Click()
'   On Error GoTo ErrorHandler
'   vStrSQL = " Select h.*,b.*, EmpName, ProductName, PackingName, UnitName from ProductionRecordHeader H " & vbCrLf _
'            + "Inner Join  ProductionRecordBody b on H.ProductionID = b.ProductionID " & vbCrLf _
'            + "Inner Join Shifts Sh on SH.EmpID = H.EmpID " & vbCrLf _
'            + "Inner Join Products Pr on PR.productiD = b.Productid " & vbCrLf _
'            + "Inner Join Packings PK on pk.PackingiD = b.PackingID" & vbCrLf _
'            + "Inner Join Units PU on PU.UnitID = b.UnitID" & vbCrLf _
'            + " where H.ProductionID = " & Val(TxtEmpID.Text)
'
'    If RsReport.State = adStateOpen Then RsReport.Close
'    RsReport.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
'
'    Set RptReportViewer.Report = New RptProductionRecord
'    RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
'    RptReportViewer.Report.SelectPrinter "Printer Driver", "Printer Name", "LPT1"
'    Dim vStrComp As String, vCompanyName As String, vAddress As String, vPhone As String
'    vStrComp = "Select CompanyName,Address,City,PhoneNo,email from Company"
'    With CN.Execute(vStrComp)
'      If .RecordCount > 0 Then
'         vCompanyName = !CompanyName
'         vAddress = !Address & IIf(IsNull(!City), "", ", " & !City)
'         vPhone = IIf(IsNull(!PhoneNo), "", "Phone # " & !PhoneNo)
'         RptReportViewer.Report.ParameterFields(1).AddCurrentValue vCompanyName
'         RptReportViewer.Report.ParameterFields(2).AddCurrentValue vAddress
'         RptReportViewer.Report.ParameterFields(3).AddCurrentValue vPhone
'      End If
'   End With
'   'RptReportViewer.Report.ParameterFields(3).AddCurrentValue CN.Execute("Select Name from Manufacturer").Fields(0).Value
'   'RptReportViewer.Report.PrintOut False
'   RptReportViewer.Show vbModal, Me
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub BtnSave_Click()
  On Error GoTo ErrorHandler
  If vIsNewRecord = False And ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsEdit = False Then
    MsgBox "You are not authorized to modify a posted record", vbCritical, "Error"
    Exit Sub
  End If
  
'  Header Validation
   If Trim(TxtEmpID.Text) = "" Then
      MsgBox "Enter Employee ID.", vbExclamation, Me.Caption
      TxtEmpID.SetFocus
      Exit Sub
   End If
   
   Dim Rs As New ADODB.Recordset
      sSql = "select * from EmpAttendance where EmpID = " & Val(TxtEmpID.Text) & " And AttendDate = '" & DtpAttendDate.DateValue & "'"
   With Rs
          Rs.Open sSql, cn, adOpenStatic, adLockReadOnly
          If .RecordCount > 0 Then
           If !AttendID <> Val(TxtAttendID.Text) Then
            MsgBox "This Employee Already done his attendance.", vbExclamation, Me.Caption
            If TxtEmpID.Enabled = True Then TxtEmpID.SetFocus
           Exit Sub
           End If
           If DtpAttendDate.DateValue = !DateOut Then
                If Format(DtpTimeIn.Value, "hh:mm") > Format(!TimeOut, "hh:mm") Then
                    MsgBox "Time In Should be Less than Time Out.", vbExclamation, Me.Caption
                    Exit Sub
                End If
            ElseIf DtpAttendDate.DateValue > !DateOut Then
                MsgBox "Date In Should be Less than Date Out.", vbExclamation, Me.Caption
                Exit Sub
            End If
          End If
   End With
   Rs.Close
   cn.BeginTrans
   sSql = "select * from EmpAttendance where EmpID = " & Val(TxtEmpID.Text) & " And AttendDate = '" & DtpAttendDate.DateValue & "'"

   With Rs
      .Open sSql, cn, adOpenStatic, adLockPessimistic
      If .RecordCount = 0 Then
         .AddNew
         !AttendID = Val(TxtAttendID.Text)
         !EmpID = Val(TxtEmpID.Text)
         !AttendDate = DtpAttendDate.DateValue
         !TimeUpdated = 0
      Else
         !TimeUpdated = 1
      End If
      !TimeIn = DtpAttendDate.DateValue & " " & Format(DtpTimeIn.Value, "hh:mm")
      !Fine = Val(TxtFine.Text)
      !UserNo = vUser
      .Update
      .Close
   End With
   cn.Execute ("Update EmpAttendance set WorkingTime =  dateDiff(Minute,timein,timeout) where EmpID = " & Val(TxtEmpID.Text) & " And AttendDate = '" & DtpAttendDate.DateValue & "' and DateOut is not null")
   cn.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If cn.Errors.Count > 0 Then cn.RollbackTrans
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
      Call SubClearFields
      BtnOpen.Enabled = True
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      BtnPrint.Enabled = False
      TxtAttendID.Text = FunGetMaxID()
      TxtEmpID.Enabled = True
      BtnEmployee.Enabled = True
      DtpAttendDate.Enabled = True
      DtpAttendDate.DateValue = Date
      If TxtEmpID.Enabled And TxtEmpID.Visible Then TxtEmpID.SetFocus
      vIsNewRecord = True
   Case Is = OpenMode
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      BtnPrint.Enabled = True
      TxtEmpID.Enabled = False
      BtnEmployee.Enabled = False
      DtpAttendDate.Enabled = False
      vIsNewRecord = False
   Case Is = changeMode
      BtnPrint.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub DtpAttendDate_Change()
FormStatus = changeMode
End Sub
Private Sub DtpTimeIn_Change()
   If DtpTimeIn.Visible = False Then Exit Sub
   If ActiveControl.Name <> DtpTimeIn.Name Then Exit Sub
   If TxtEmpID.Text = "" Then Exit Sub
   CalculateFine
   If BtnSave.Enabled = False Then FormStatus = changeMode

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
         keybd_event 9, 1, 1, 1
            KeyCode = 0
   ElseIf KeyCode = vbKeyEscape Then
      'Call SubClearDetailArea: TxtProductID.SetFocus
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
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtEmpID.Name: If FunSelectEmployee(ssFunctionKey, False) = True Then DtpAttendDate.SetFocus
      End Select
'   ElseIf ActiveControl.Name = TxtProductID.Name Then
'      If KeyCode = vbKeyDown Then
'         Grid.SetFocus
'      ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtProductID.Name Then
'         KeyCode = 0
'         TxtDescription.SetFocus
      End If
  
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = changeMode
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Employee Attendace IN"
   
   DtpAttendDate.DateValue = IIf(Format(Now, "hh") >= Val(ObjRegistry.HourDifference), Date, DateAdd("d", -1, Date))
   
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = cn.Execute("Select isnull(max(AttendID),0)+1 from EmpAttendance").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

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
   TxtFinePerMinute.Text = ""
   TxtFine.Text = ""
   DtpAttendDate.DateValue = Date
   DtpTimeIn.Value = Time
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Are you sure to close without save?", vbQuestion + vbApplicationModal + vbYesNo, "Alert") = vbNo Then
         Cancel = 1
      End If
   Else
    Dim frmObj As Object
    For Each frmObj In Forms
        Set frmObj = Nothing
    Next
    Set FrmEmployeeAttendanceIn = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub GetEmployeeAttendace()
   On Error GoTo ErrorHandler
   sSql = "select EA.*, Emp.EmpName, OfficeTimeIn, FinePerMinute FROM EmpAttendance EA inner join Employees Emp on Emp.EmpID = EA.EmpID  where EA.EmpID = " & Val(TxtEmpID.Text) & " and AttendDate ='" & DtpAttendDate.DateValue & "'"
   With cn.Execute(sSql)
   If .RecordCount > 0 Then
      If Not .BOF Then
          TxtAttendID.Text = !AttendID
          DtpAttendDate.DateValue = !AttendDate
          DtpTimeIn.Value = !TimeIn
          DtpOfficeTimeIn.Value = IIf(IsNull(!OfficeTimeIn), Time, !OfficeTimeIn)
          TxtFinePerMinute.Text = IIf(IsNull(!FinePerMinute), "", !FinePerMinute)
          TxtFine.Text = IIf(IsNull(!Fine), "", !Fine)
          TxtEmpID.Text = !EmpID
          TxtEmpName.Text = !EmpName
      End If
    End If
      .Close
   End With
   DtpTimeIn.SetFocus
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtEmpID_Change()
   If TxtEmpID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtEmpID.Name Then Exit Sub
   If TxtEmpName.Text <> "" Then TxtEmpName.Text = ""
End Sub

Private Sub TxtEmpID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtEmpID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtEmpName.Text <> "" Then Exit Sub
   If Trim(TxtEmpID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectEmployee(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectEmployee(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Function FunSelectEmployee(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchEmployee.Show vbModal, Me
        If SchEmployee.ParaOutEmployeeID = "" Then FunSelectEmployee = False: Exit Function
        TxtEmpID.Text = SchEmployee.ParaOutEmployeeID
    End If
    '---------------------------
    vStrSQL = " Select EmpName, OfficeTimeIn, FinePerMinute  FROM Employees where EmpID = " & Val(TxtEmpID.Text)
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtEmpName.Text = !EmpName
          FunSelectEmployee = True
          DtpAttendDate.DateValue = Date
          DtpTimeIn.Value = Time
          DtpOfficeTimeIn.Value = !OfficeTimeIn
          TxtFinePerMinute.Text = IIf(IsNull(!FinePerMinute), 0, !FinePerMinute)
          CalculateFine
          .Close
          If BtnSave.Enabled = False Then FormStatus = changeMode
          Exit Function
      Else
          FunSelectEmployee = False
          .Close
          TxtEmpID.Text = ""
          TxtEmpName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = changeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
Private Sub CalculateFine()
On Error GoTo ErrorHandler
   vMinute = DateDiff("n", Format(DtpOfficeTimeIn.Value, "hh:mm"), Format(DtpTimeIn.Value, "hh:mm"))
   vMinute = vMinute - ObjRegistry.EmployeeLateRelaxTime
   If vMinute > 0 Then
      vMinute = vMinute + ObjRegistry.EmployeeLateRelaxTime
      TxtFine.Text = vMinute * Val(TxtFinePerMinute.Text)
   Else
      TxtFine.Text = 0
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
