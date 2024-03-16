VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form RptMemberList 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptMemberList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8430
      TabIndex        =   4
      Top             =   6353
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Close"
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
      MICON           =   "RptMemberList.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      Height          =   420
      Left            =   5655
      TabIndex        =   2
      Top             =   6353
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Pre&view"
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
      MICON           =   "RptMemberList.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7035
      TabIndex        =   3
      Top             =   6353
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "&Print"
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
      MICON           =   "RptMemberList.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtFromMemberID 
      Height          =   315
      Left            =   6230
      TabIndex        =   0
      Top             =   4613
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   6
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
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtToMemberID 
      Height          =   315
      Left            =   7865
      TabIndex        =   1
      Top             =   4613
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   6
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
      IntegralPoint   =   3
   End
   Begin JeweledBut.JeweledButton BtnFromMember 
      Height          =   330
      Left            =   7125
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4613
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
      MICON           =   "RptMemberList.frx":0F1E
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnToMember 
      Height          =   330
      Left            =   8760
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4613
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
      MICON           =   "RptMemberList.frx":0F3A
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Member ID"
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
      Left            =   7865
      TabIndex        =   7
      Top             =   4358
      Width           =   1215
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Member List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   6
      Top             =   270
      Width           =   1635
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Member ID"
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
      Left            =   6230
      TabIndex        =   5
      Top             =   4373
      Width           =   1395
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove this Row"
      End
   End
End
Attribute VB_Name = "RptMemberList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flag As Boolean
Dim Rs As New ADODB.Recordset
Dim sSql As String
Dim VStrSQL As String

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnFromMember_Click()
   If FunSelectFromMember(ssButton, True) = True Then
      TxtToMemberID.SetFocus
   Else
      TxtFromMemberID.SetFocus
   End If
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
       RptReportViewer.Caption = "Member List Report"
       RptReportViewer.Show vbModal
   End If
End Sub

Private Sub BtnPrint_Click()
   If SetReport Then RptReportViewer.Report.PrintOut False
End Sub

Private Sub BtnToMember_Click()
   If FunSelectToMember(ssButton, True) = True Then
      BtnPreview.SetFocus
   Else
      TxtToMemberID.SetFocus
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
         Case vbKeyV
            If BtnPreview.Enabled Then BtnPreview_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtFromMemberID.Name: If FunSelectFromMember(ssFunctionKey, True) = True Then TxtToMemberID.SetFocus
         Case TxtToMemberID.Name: If FunSelectToMember(ssFunctionKey, True) = True Then BtnPreview.SetFocus
      End Select
   End If
   Exit Sub
ErrorHandler:
    Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Members List"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   Dim frmObj As Object
   For Each frmObj In Forms
       Set frmObj = Nothing
   Next
   Set RptMemberList = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Function SetReport() As Boolean
   On Error GoTo ErrorHandler
   Dim RsReport As New ADODB.Recordset
   SetReport = False
   Dim vWhere As String
   vWhere = IIf(Trim(TxtFromMemberID.Text) = "", "", IIf(Trim(TxtToMemberID.Text) = "", " and MemberID >= " & Val(TxtFromMemberID.Text), " and MemberID Between " & Val(TxtFromMemberID.Text) & " and " & Val(TxtToMemberID.Text)))
   
   VStrSQL = "Select MemberID, MemberName, isnull(Address,'') + isnull(' ('+ City + ')','') as Address, isnull(' ' + Phone1,'') + isnull(' ' + Phone2,'') + isnull(' ' + Mobile,'') as ContactNo from Members Where 1=1 " & vbCrLf _
   + vWhere
   Me.MousePointer = vbHourglass
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open VStrSQL, CN, adOpenStatic, adLockReadOnly
   Set RptReportViewer.Report = New CRptMemberList
   If RsReport.BOF Then
       MsgBox "No record exists.", vbInformation, Me.Caption
       Me.MousePointer = vbDefault
       Exit Function
   End If
   RptReportViewer.Report.ReportTitle = "Member List"
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ObjRegistry.CompanyAddress = "", "", ObjRegistry.CompanyAddress) & IIf(ObjRegistry.CompanyCity = "", "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", ".", " Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   RptReportViewer.Report.PaperOrientation = crLandscape
   SetReport = True
   Me.MousePointer = vbDefault
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectToMember(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchMember.Show vbModal, Me
      If SchMember.ParaOutMemberID = "" Then FunSelectToMember = False: Exit Function
      TxtToMemberID.Text = SchMember.ParaOutMemberID
   End If
    '---------------------------
    If Trim(TxtToMemberID.Text) = "" Then Exit Function
    If TxtToMemberID.Text = "" Then FunSelectToMember = False: Exit Function
    VStrSQL = " SELECT Memberid, MemberName" & vbCrLf _
           + " from Members " & vbCrLf _
           + " where MemberID = " & Val(TxtToMemberID.Text)
  
   With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         FunSelectToMember = True
         .Close
         Exit Function
      Else
         FunSelectToMember = False
         .Close
         MsgBox "Invalid Member ID.", vbOKOnly, "Alert"
         TxtToMemberID.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Function FunSelectFromMember(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim VStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchMember.Show vbModal, Me
      If SchMember.ParaOutMemberID = "" Then FunSelectFromMember = False: Exit Function
      TxtFromMemberID.Text = SchMember.ParaOutMemberID
   End If
    '---------------------------
    If Trim(TxtFromMemberID.Text) = "" Then Exit Function
    If TxtFromMemberID.Text = "" Then FunSelectFromMember = False: Exit Function
    VStrSQL = " SELECT Memberid, MemberName" & vbCrLf _
           + " from Members " & vbCrLf _
           + " where MemberID = " & Val(TxtFromMemberID.Text)
  
   With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
         FunSelectFromMember = True
         .Close
         Exit Function
      Else
         FunSelectFromMember = False
         .Close
         MsgBox "Invalid Member ID.", vbOKOnly, "Alert"
         TxtFromMemberID.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtToMemberID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtToMemberID.Text) = "" Then Exit Sub
   vTemp = Not FunSelectToMember(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectToMember(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtFromMemberID_Validate(Cancel As Boolean)
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtFromMemberID.Text) = "" Then Exit Sub
   vTemp = Not FunSelectFromMember(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectFromMember(ssValidate, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
