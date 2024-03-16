VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmMeterReadings 
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
   Begin JeweledBut.JeweledButton BtnDelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8363
      TabIndex        =   9
      Top             =   7740
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
      MICON           =   "FrmMeterReadings.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   7043
      TabIndex        =   6
      Top             =   7740
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
      MICON           =   "FrmMeterReadings.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   4403
      TabIndex        =   8
      Top             =   7740
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
      MICON           =   "FrmMeterReadings.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9683
      TabIndex        =   10
      Top             =   7740
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
      MICON           =   "FrmMeterReadings.frx":0054
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5723
      TabIndex        =   7
      Top             =   7740
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
      MICON           =   "FrmMeterReadings.frx":0070
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpReadingDate 
      Height          =   315
      Left            =   7328
      TabIndex        =   1
      Top             =   3210
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
   Begin SITextBox.Txt TxtReadingID 
      Height          =   315
      Left            =   6038
      TabIndex        =   0
      Top             =   3210
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   9
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
   End
   Begin SITextBox.Txt TxtEmpName 
      Height          =   315
      Left            =   7133
      TabIndex        =   14
      Top             =   4080
      Width           =   2175
      _ExtentX        =   3836
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
   Begin JeweledBut.JeweledButton BtnEmp 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6773
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4080
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
      MICON           =   "FrmMeterReadings.frx":008C
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtEmpID 
      Height          =   315
      Left            =   6038
      TabIndex        =   2
      Top             =   4080
      Width           =   735
      _ExtentX        =   1296
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
   Begin SITextBox.Txt TxtShiftName 
      Height          =   315
      Left            =   7133
      TabIndex        =   18
      Top             =   4890
      Width           =   2175
      _ExtentX        =   3836
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
   Begin JeweledBut.JeweledButton BtnShift 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6773
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4890
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
      MICON           =   "FrmMeterReadings.frx":00A8
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtShiftID 
      Height          =   315
      Left            =   6038
      TabIndex        =   3
      Top             =   4890
      Width           =   735
      _ExtentX        =   1296
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
   Begin SITextBox.Txt TxtSaleCounterName 
      Height          =   315
      Left            =   7133
      TabIndex        =   22
      Top             =   5745
      Width           =   2175
      _ExtentX        =   3836
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
   Begin JeweledBut.JeweledButton BtnSaleCounter 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   6773
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5745
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
      MICON           =   "FrmMeterReadings.frx":00C4
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtEndReading 
      Height          =   315
      Left            =   7703
      TabIndex        =   5
      Top             =   6555
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
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
      Masked          =   1
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtDifference 
      Height          =   315
      Left            =   9278
      TabIndex        =   27
      Top             =   6555
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
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
      Masked          =   1
      IntegralPoint   =   7
   End
   Begin SITextBox.Txt TxtSaleCounterID 
      Height          =   315
      Left            =   6038
      TabIndex        =   4
      Top             =   5745
      Width           =   735
      _ExtentX        =   1296
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
   Begin SITextBox.Txt TxtStartReading 
      Height          =   315
      Left            =   6038
      TabIndex        =   29
      Top             =   6555
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
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
      Masked          =   1
      IntegralPoint   =   7
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Start Reading"
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
      Left            =   6038
      TabIndex        =   30
      Top             =   6330
      Width           =   1185
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Differnce"
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
      Left            =   9278
      TabIndex        =   28
      Top             =   6330
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "End Reading"
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
      Left            =   7703
      TabIndex        =   26
      Top             =   6330
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Counter Name"
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
      Left            =   7628
      TabIndex        =   25
      Top             =   5535
      Width           =   1650
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Counter  ID"
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
      Left            =   6038
      TabIndex        =   24
      Top             =   5535
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift Name"
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
      Left            =   7133
      TabIndex        =   21
      Top             =   4680
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift ID"
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
      Left            =   6038
      TabIndex        =   20
      Top             =   4680
      Width           =   660
   End
   Begin VB.Label LblStoreName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Name"
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
      Left            =   7133
      TabIndex        =   17
      Top             =   3870
      Width           =   915
   End
   Begin VB.Label LblStoreID 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID"
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
      Left            =   6038
      TabIndex        =   16
      Top             =   3870
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Reading ID"
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
      Left            =   6038
      TabIndex        =   13
      Top             =   2970
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reading Date"
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
      Left            =   7328
      TabIndex        =   12
      Top             =   2970
      Width           =   1185
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   360
   End
   Begin VB.Label LblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Meter Readings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   11
      Top             =   270
      Width           =   2865
   End
End
Attribute VB_Name = "FrmMeterReadings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCounter As Integer
Dim Flag As Boolean
Dim sSql As String
Dim vStrSQL As String
Dim vMode As FormMode
Dim vIsNewRecord As Boolean
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
  CN.BeginTrans
  CN.Execute "Delete from MeterReadings Where ReadingID = " & Val(TxtReadingID.Text)
  CN.CommitTrans
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  If CN.Errors.Count > 0 Then CN.RollbackTrans
  Call ShowErrorMessage
End Sub

Private Sub GetMeterReading()
   On Error GoTo ErrorHandler
   sSql = "Select * From MeterReadings r inner join employees e on e.EmpID = e.EmpID inner join Shifts s on s.ShiftID = r.ShiftID inner join SaleCounters sc on sc.SaleCounterID = r.SaleCounterID where ReadingID = " & Val(TxtReadingID.Text)
   With CN.Execute(sSql)
      If Not .BOF Then
         DtpReadingDate.DateValue = !ReadingDate
         TxtEmpID.Text = !EmpID
         TxtEmpName.Text = !EmpName
         TxtShiftID.Text = !ShiftID
         TxtShiftName.Text = !ShiftName
         TxtSaleCounterID.Text = !SaleCounterID
         TxtSaleCounterName.Text = !SaleCounterName
         TxtStartReading.Text = !StartReading
         TxtEndReading.Text = !EndReading
      End If
      .Close
   End With
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOpen_Click()
   SchMeterReading.Show vbModal
   If SchMeterReading.ParaOutID <> Empty Then
      TxtReadingID.Text = SchMeterReading.ParaOutID
      GetMeterReading
   End If
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
   If vIsNewRecord Then
      If CN.Execute("Select * from MeterReadings where ReadingID = " & Val(TxtReadingID.Text)).RecordCount > 0 Then
         MsgBox "This voucher already exists. A new voucher No. has been generated. Please try again", vbCritical, "Alert"
         TxtReadingID.Text = FunGetMaxID
         Exit Sub
      End If
      If CN.Execute("Select * from MeterReadings where ReadingDate = '" & DtpReadingDate.DateValue & "' and EmpID = " & Val(TxtReadingID.Text) & " and ShiftID = " & Val(TxtShiftID.Text) & " and SaleCounterID = " & TxtSaleCounterID.Text).RecordCount > 0 Then
         MsgBox "This Meter Reading Already Entered. Please try again", vbCritical, "Alert"
         TxtReadingID.Text = FunGetMaxID
         Exit Sub
      End If
   End If
   If Trim(TxtEmpID.Text) = "" Then
      MsgBox "Enter Emp ID.", vbExclamation, Me.Caption
      TxtEmpID.SetFocus
      Exit Sub
   End If
   If Trim(TxtShiftID.Text) = "" Then
      MsgBox "Enter Shift ID.", vbExclamation, Me.Caption
      TxtShiftID.SetFocus
      Exit Sub
   End If
   
   If Trim(TxtSaleCounterID.Text) = "" Then
      MsgBox "Enter Sale Counter ID.", vbExclamation, Me.Caption
      TxtSaleCounterID.SetFocus
      Exit Sub
   End If
         
  
  'Saving record
   CN.BeginTrans
   sSql = "Select * From MeterReadings Where ReadingID =" & Val(TxtReadingID.Text)
   Dim Rs As New ADODB.Recordset
   With Rs
      .Open sSql, CN, adOpenDynamic, adLockOptimistic
      If .BOF Then
         .AddNew
         !ReadingID = Val(TxtReadingID.Text)
      End If
      !ReadingDate = DtpReadingDate.DateValue
      !EmpID = TxtEmpID.Text
      !ShiftID = TxtShiftID.Text
      !SaleCounterID = TxtSaleCounterID.Text
      !StartReading = Val(TxtStartReading.Text)
      !EndReading = Val(TxtEndReading.Text)
      !UserNo = vUser
      .Update
      .Close
   End With
   CN.CommitTrans
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   If CN.Errors.Count > 0 Then CN.RollbackTrans
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
      TxtReadingID.Text = FunGetMaxID
      If DtpReadingDate.Enabled And DtpReadingDate.Visible Then DtpReadingDate.SetFocus
      vIsNewRecord = True
    Case Is = OpenMode
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
      vIsNewRecord = False
    Case Is = changeMode
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = True
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub DtpReadingDate_Change()
'    If DtpReadingDate.Visible = False Then Exit Sub
'    If Me.ActiveControl.Name <> DtpReadingDate.Name Then Exit Sub
'    TxtReadingID.Text = FunGetMaxID
'    If DtpReadingDate.Enabled And DtpReadingDate.Visible Then FormStatus = ChangeMode
End Sub

Private Sub DtpReadingDate_Click()
'    If DtpReadingDate.Visible = False Then Exit Sub
'    If Me.ActiveControl.Name <> DtpReadingDate.Name Then Exit Sub
'    TxtReadingID.Text = FunGetMaxID
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
  ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtEmpID.Name: If FunSelectEmployee(ssFunctionKey, False) = True Then TxtShiftID.SetFocus
         Case TxtShiftID.Name: If FunSelectShift(ssFunctionKey, False) = True Then TxtSaleCounterID.SetFocus
         Case TxtSaleCounterID.Name: If FunSelectSaleCounter(ssFunctionKey, False) = True Then TxtEndReading.SetFocus
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
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'   Select Case ActiveControl.Name
'   Case TxtUnderQty.Name, TxtProductID.Name
''      Call NonNumeric(KeyAscii, ActiveControl, False)
'   End Select
   If BtnSave.Enabled Then Exit Sub
   If UCase(Me.ActiveControl.Name) Like "TXT*" Then FormStatus = changeMode
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Meter Readings"
'   With CN.Execute("select * from Registry")
'      If .RecordCount > 0 Then
'         TxtStoreID.Text = IIf(IsNull(!StoreID), "", !StoreID)
'         FunSelectStore ssValidate, True
'         TxtStoreID.Visible = !StoreVisible
'         BtnStore.Visible = !StoreVisible
'         TxtStoreName.Visible = !StoreVisible
'         LblStoreID.Visible = !StoreVisible
'         LblStoreName.Visible = !StoreVisible
'      End If
'      .Close
'   End With
   FormStatus = NewMode
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As Long
  On Error GoTo ErrorHandler
  FunGetMaxID = CN.Execute("Select isnull(max(ReadingID),0) from MeterReadings").Fields(0) + 1
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub SubClearFields()
  On Error GoTo ErrorHandler
  Dim ctl As Control
  For Each ctl In Me.Controls
    If TypeOf ctl Is TextBox Then
      ctl.Text = ""
    'ElseIf TypeOf ctl Is ComboBox Then
    ElseIf TypeOf ctl Is SITextBox.txt Then
      If ctl.Tag = "" Then ctl.Text = ""
    End If
  Next
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
    Set FrmMeterReadings = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
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
    vStrSQL = " Select * FROM Employees where EmpID = " & Val(TxtEmpID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtEmpName.Text = !EmpName
          FunSelectEmployee = True
          .Close
          If BtnSave.Enabled = False And BtnSave.Visible = True Then FormStatus = changeMode
          Exit Function
      Else
          FunSelectEmployee = False
          .Close
          TxtEmpID.Text = ""
          TxtEmpName.Text = ""
          If BtnSave.Enabled = False And BtnSave.Visible = True Then FormStatus = changeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtEmpID_Change()
   On Error GoTo ErrorHandler
   If TxtEmpID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtEmpID.Name Then Exit Sub
   If TxtEmpName.Text <> "" Then TxtEmpName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtEmpID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtEmpID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtEmpName.Text <> "" Then Exit Sub
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

Private Sub BtnEmp_Click()
   On Error GoTo ErrorHandler
   If FunSelectEmployee(ssButton, False) = True Then
      TxtShiftID.SetFocus
   Else
      TxtEmpID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectShift(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchShift.Show vbModal, Me
        If SchShift.ParaOutShiftID = "" Then FunSelectShift = False: Exit Function
        TxtShiftID.Text = SchShift.ParaOutShiftID
    End If
    '---------------------------
    vStrSQL = " Select * FROM Shifts where ShiftID=" & Val(TxtShiftID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtShiftName.Text = !ShiftName
          FunSelectShift = True
          .Close
          If BtnSave.Enabled = False And BtnSave.Visible = True Then FormStatus = changeMode
          Exit Function
      Else
          FunSelectShift = False
          .Close
          TxtShiftID.Text = ""
          TxtShiftName.Text = ""
          If BtnSave.Enabled = False And BtnSave.Visible = True Then FormStatus = changeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtEndReading_Change()
   On Error GoTo ErrorHandler
   If Val(TxtStartReading.Text) = 0 Then
      TxtDifference.Text = 0
   Else
      TxtDifference.Text = Val(TxtEndReading.Text) - Val(TxtStartReading.Text)
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtEndReading_GotFocus()
   On Error GoTo ErrorHandler
   If ActiveControl.Name <> TxtEndReading.Name Then Exit Sub
   TxtStartReading.Text = CN.Execute("select dbo.FunStartReading('" & DtpReadingDate.DateValue & "','" & TxtEmpID.Text & "'," & TxtShiftID.Text & "," & TxtSaleCounterID.Text & ")").Fields(0).Value
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtShiftID_Change()
   On Error GoTo ErrorHandler
   If TxtShiftID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtShiftID.Name Then Exit Sub
   If TxtShiftName.Text <> "" Then TxtShiftName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtShiftID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtShiftID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtShiftName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectShift(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectShift(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnShift_Click()
   On Error GoTo ErrorHandler
   If FunSelectShift(ssButton, False) = True Then
      TxtSaleCounterID.SetFocus
   Else
      TxtShiftID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectSaleCounter(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchSaleCounter.Show vbModal, Me
        If SchSaleCounter.ParaOutSaleCounterID = "" Then FunSelectSaleCounter = False: Exit Function
        TxtSaleCounterID.Text = SchSaleCounter.ParaOutSaleCounterID
    End If
    '---------------------------
    vStrSQL = " Select * FROM SaleCounters where SaleCounterID=" & Val(TxtSaleCounterID.Text)
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtSaleCounterName.Text = !SaleCounterName
          FunSelectSaleCounter = True
          .Close
          If BtnSave.Enabled = False And BtnSave.Visible = True Then FormStatus = changeMode
          Exit Function
      Else
          FunSelectSaleCounter = False
          .Close
          TxtSaleCounterID.Text = ""
          TxtSaleCounterName.Text = ""
          If BtnSave.Enabled = False And BtnSave.Visible = True Then FormStatus = changeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtSaleCounterID_Change()
   On Error GoTo ErrorHandler
   If TxtSaleCounterID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtSaleCounterID.Name Then Exit Sub
   If TxtSaleCounterName.Text <> "" Then TxtSaleCounterName.Text = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSaleCounterID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtSaleCounterID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtSaleCounterName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectSaleCounter(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectSaleCounter(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSaleCounter_Click()
   On Error GoTo ErrorHandler
   If FunSelectSaleCounter(ssButton, False) = True Then
      TxtEndReading.SetFocus
   Else
      TxtSaleCounterID.SetFocus
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

