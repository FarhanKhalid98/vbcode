VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptBalanceSheet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptBalanceSheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkExclude 
      BackColor       =   &H00B98A03&
      Caption         =   "Exclude Accounts Having Zero Balance."
      Height          =   255
      Left            =   7099
      TabIndex        =   6
      Top             =   7478
      Visible         =   0   'False
      Width           =   3285
   End
   Begin JeweledBut.JeweledButton CmdPreview 
      Height          =   420
      Left            =   5479
      TabIndex        =   3
      Top             =   5618
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Preview"
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
      MICON           =   "RptBalanceSheet.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdPrint 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   6829
      TabIndex        =   4
      Top             =   5618
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Print"
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
      MICON           =   "RptBalanceSheet.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton CmdClose 
      Height          =   420
      Left            =   8179
      TabIndex        =   5
      Top             =   5618
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Close"
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
      MICON           =   "RptBalanceSheet.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5997
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   3623
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "RptBalanceSheet.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtOrganizationID 
      Height          =   315
      Left            =   4977
      TabIndex        =   0
      Top             =   3623
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   2
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
   Begin SITextBox.Txt TxtOrganizatonName 
      Height          =   315
      Left            =   6357
      TabIndex        =   9
      Tag             =   "nc"
      Top             =   3623
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpFrom 
      Height          =   315
      Left            =   5944
      TabIndex        =   1
      Top             =   4598
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
   Begin SSCalendarWidgets_A.SSDateCombo DtpTo 
      Height          =   315
      Left            =   7669
      TabIndex        =   2
      Top             =   4598
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
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
      Left            =   5944
      TabIndex        =   13
      Top             =   4373
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
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
      Left            =   7699
      TabIndex        =   12
      Top             =   4373
      Width           =   705
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization ID"
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
      Left            =   4984
      TabIndex        =   11
      Top             =   3398
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
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
      Left            =   6364
      TabIndex        =   10
      Top             =   3398
      Width           =   1620
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Sheet"
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
      TabIndex        =   7
      Top             =   270
      Width           =   1890
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "RptBalanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vStrSQL As String

Private Sub BtnGroup_Click()
   If FunSelectOrganizaton(ssButton, False) = True Then
      DtpFrom.SetFocus
   Else
      TxtOrganizationID.SetFocus
   End If
End Sub

Private Sub CmdClose_Click()
   Unload Me
End Sub

Private Sub CmdPreview_Click()
  On Error GoTo ErrorHandler
  If FunRefreshData = False Then Exit Sub
  If Rs.RecordCount = 0 Then
    MsgBox "No record found", vbInformation, "Information"
    Exit Sub
  Else
    Call SetCrystalReport
    RptReportViewer.Show vbModal, Me
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub CmdPrint_Click()
  On Error GoTo ErrorHandler
  If FunRefreshData = False Then Exit Sub
  If Rs.RecordCount = 0 Then
    MsgBox "No record found", vbInformation, "Information"
    Exit Sub
  Else
    Call SetCrystalReport
    RptReportViewer.Report.PrintOut
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtOrganizationID.Name: If FunSelectOrganizaton(ssFunctionKey, True) = True Then DtpFrom.SetFocus
      End Select
   End If
End Sub

Private Function FunRefreshData() As Boolean
   On Error GoTo ErrorHandler
   
   Dim vSQL As String, vWhere  As String
   vWhere = " and (AccountsBalances.Debit > 0 OR AccountsBalances.Credit > 0 or AccountsBalances.OpeningBal > 0 or AccountsBalances.Bal > 0 )" & IIf(Trim(TxtOrganizationID.Text) = "", "", " And AccountsBalances.OrganizationID = " & TxtOrganizationID.Text)
   
   vSQL = "EXECUTE SPAccountsBalancesNew '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'"
   CN.Execute "EXECUTE SPAccountsBalancesNew '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'"
   
  
   
   'First Delete Head Account and Purchases
   vSQL = " Delete From AccountsBalances " & vbCrLf _
      + " from AccountsBalances a inner join Chartofaccounts c on c.AccountNo = a.AccountNo" & vbCrLf _
      + " where isDetailed = 0 or a.AccountNo in('112','113','21')"
   Set Rs = CN.Execute(vSQL)
   
   'Second Insert Closing Stock
   CN.Execute "EXECUTE SPClosingStock '" & DtpFrom.DateValue & "','" & DtpTo.DateValue & "'"
  
   vSQL = " Delete From AccountsBalances " & vbCrLf _
      + " from AccountsBalances a inner join Chartofaccounts c on c.AccountNo = a.AccountNo" & vbCrLf _
      + " where a.AccountNo in ( Select Value from Defaultvalues where [Key] = 'Purchase' or [Key] = 'Cost of Goods Sold')"
   Set Rs = CN.Execute(vSQL)
  
   'Third Insert Account of Receiveables and Payables
   vSQL = " insert into AccountsBalances  " & vbCrLf _
      + " Select '113' Accountno, Sum(OpeningDebit),Sum(OpeningCredit),Abs(Sum(OpeningDebit)-Sum(OpeningCredit))," & vbCrLf _
      + " case when Sum(OpeningDebit)-Sum(OpeningCredit) >= 0 then 'Dr' else 'Cr' end," & vbCrLf _
      + " Sum(Debit),Sum(Credit),Abs(Sum(OpeningDebit)-Sum(OpeningCredit)+Sum(Debit)-Sum(Credit)), case when Sum(OpeningDebit)-Sum(OpeningCredit)+Sum(Debit)-Sum(Credit) >= 0 then 'Dr' else 'Cr' end, OrganizationID" & vbCrLf _
      + " FROM AccountsBalances " & vbCrLf _
      + " where AccountNo like '6%' and BalType = 'Dr'" & vbCrLf _
      + " Group By OrganizationID" & vbCrLf _
      + " union all" & vbCrLf _
      + " Select '21' Accountno, Sum(OpeningDebit),Sum(OpeningCredit),Abs(Sum(OpeningDebit)-Sum(OpeningCredit))," & vbCrLf _
      + " case when Sum(OpeningDebit)-Sum(OpeningCredit) >= 0 then 'Dr' else 'Cr' end," & vbCrLf _
      + " Sum(Debit),Sum(Credit),Abs(Sum(OpeningDebit)-Sum(OpeningCredit)+Sum(Debit)-Sum(Credit)), case when Sum(OpeningDebit)-Sum(OpeningCredit)+Sum(Debit)-Sum(Credit) >= 0 then 'Dr' else 'Cr' end, OrganizationID" & vbCrLf _
      + " FROM AccountsBalances " & vbCrLf _
      + " where AccountNo like '6%' and BalType = 'Cr'" & vbCrLf _
      + " Group By OrganizationID"
  Set Rs = CN.Execute(vSQL)
  
  'Fourth Insert Head Accounts
  vSQL = " INSERT INTO AccountsBalances" & vbCrLf _
      + "      Select c.AccountNo, Sum(a.OpeningDebit),Sum(a.OpeningCredit),Abs(Sum(a.OpeningDebit)-Sum(a.OpeningCredit))," & vbCrLf _
      + "      case when Sum(a.OpeningDebit)-Sum(a.OpeningCredit) >= 0 then 'Dr' else 'Cr' end," & vbCrLf _
      + "      Sum(a.Debit),Sum(a.Credit),Abs(Sum(a.OpeningDebit)-Sum(a.OpeningCredit)+Sum(a.Debit)-Sum(a.Credit)), case when Sum(a.OpeningDebit)-Sum(a.OpeningCredit)+Sum(a.Debit)-Sum(a.Credit) >= 0 then 'Dr' else 'Cr' end, OrganizationID" & vbCrLf _
      + "      FROM ChartOfAccounts c INNER JOIN AccountsBalances a ON c.AccountNo = LEFT(a.AccountNo,LEN(c.AccountNo))" & vbCrLf _
      + "      WHERE IsDetailed = 0" & vbCrLf _
      + "      GROUP BY c.AccountNo, OrganizationID"
  Set Rs = CN.Execute(vSQL)
  
'''  'Start Query to Update Profit and Loss Bal
  vSQL = "Update AccountsBalances Set OpeningDebit = 0, OpeningCredit = 0, OPeningBalType = PLBalType, Debit = 0, Credit = 0,  Bal = (case when BalType = 'Cr' and PLBalType = 'Cr'  then  PLBal+Bal else PLBal-Bal end), BalType = PLBalType " & vbCrLf _
      + " --Select AB.AccountNO, SQ_PL.AccountNo, SQ_PL.OrganizationID, PLBal, PLBalType" & vbCrLf _
      + " From AccountsBalances AB" & vbCrLf _
      + " Inner Join" & vbCrLf _
      + " (" & vbCrLf _
      + "      Select AccountNo, OrganizationID, ABS(Sum(DrBal)-Sum(CrBal)) PLBal, case when Sum(DrBal)-Sum(CrBal) >= 0 then 'Cr' else 'Dr' end PLBalType From" & vbCrLf _
      + "      (" & vbCrLf _
      + "      Select   dbo.DefaultValue('Loss') AccountNo, AB.OrganizationID, sum(Bal) CrBal, 0 DrBal from  Chartofaccounts CofA" & vbCrLf _
      + "      Left Outer Join AccountsBalances AB On AB.AccountNo= CofA.AccountNo" & vbCrLf _
      + "      Where not(CofA.Accountno like '6%' or CofA.Accountno like '5%' or CofA.Accountno like '4%')" & vbCrLf _
      + "      and isdetailed <> 0 and baltype = 'Cr'" & vbCrLf _
      + "      group by  AB.OrganizationID" & vbCrLf _
      + "      Union All" & vbCrLf _
      + "      Select  dbo.DefaultValue('Loss') AccountNo, AB.OrganizationID, 0 CrBal, sum(Bal) DrBal from  Chartofaccounts CofA" & vbCrLf _
      + "      Left Outer Join AccountsBalances AB On AB.AccountNo= CofA.AccountNo" & vbCrLf _
      + "      Where not(CofA.Accountno like '6%' or CofA.Accountno like '5%' or CofA.Accountno like '4%')" & vbCrLf _
      + "      and isdetailed <> 0 and baltype = 'Dr'" & vbCrLf _
      + "      group by  AB.OrganizationID" & vbCrLf _
      + "      )PLBaL" & vbCrLf _
      + "      Group by AccountNo, OrganizationID" & vbCrLf _
      + " )SQ_PL" & vbCrLf _
      + " On SQ_PL.AccountNO = AB.AccountNO and isnull(SQ_PL.OrganizationID,'') = isnull(AB.OrganizationID,'')"
  Set Rs = CN.Execute(vSQL)
  vSQL = "Update AccountsBalances Set AccountNo = dbo.DefaultValue('Profit') Where AccountNo = dbo.DefaultValue('Loss') and BalType = 'Cr'"
  Set Rs = CN.Execute(vSQL)
  ''''''''''''''''  End Query to Update Profit and Loss Bal '''''''''''''''''''''
  
  vSQL = "Select cast(cofA.AccountNo as varchar(10)) as AccountNo, cofA.AccountName, cofA.AccountType, cofA.AccountDepth, cofA.isDetailed, AB.OrganizationID, OrganizationName, Debit, Credit, Bal, BalType, case  When isDetailed = 0 then 'General' Else 'Detail' End Nature  from Chartofaccounts CofA " & vbCrLf _
        + " Left Outer Join AccountsBalances AB On AB.AccountNo= CofA.AccountNo" & vbCrLf _
        + " left Outer Join Organizations O On O.OrganizationID = AB.OrganizationID" & vbCrLf _
        + " Where not(CofA.Accountno like '6%' or CofA.Accountno like '5%' or CofA.Accountno like '4%')" & vbCrLf _
        + IIf(Trim(TxtOrganizationID.Text) = "", "", " And O.OrganizationID = " & TxtOrganizationID.Text) & vbCrLf _
        + " And Bal <> 0 " & vbCrLf _
        + "order by CofA.Accountno"
  Set Rs = CN.Execute(vSQL)
  
  FunRefreshData = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
  FunRefreshData = False
End Function

Private Sub SetCrystalReport()
   On Error GoTo ErrorHandler
   Set RptReportViewer.Report = New CrptBalanceSheet
  
   RptReportViewer.Report.ReportTitle = "Balance Sheet"
   RptReportViewer.Report.Database.SetDataSource Rs, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue "From : " & Format(DtpFrom.DateValue, "dd/MM/yyyy") & ",   To : " & Format(DtpTo.DateValue, "dd/MM/yyyy")
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue IIf(ChkExclude.Value = 1, "Excluding", "Including") & " accounts having zero balance"
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(6).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   RptReportViewer.Report.PaperOrientation = crPortrait
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Balance Sheet"
   DtpFrom.DateValue = Date - 30
   DtpTo.DateValue = Date
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtOrganizationID_Change()
   If TxtOrganizationID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   If TxtOrganizatonName.Text <> "" Then TxtOrganizatonName.Text = ""
End Sub

Private Sub TxtOrganizationID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtOrganizationID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If Trim(TxtOrganizationID.Text) = "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectOrganizaton(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectOrganizaton(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectOrganizaton(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchOrganization.Show vbModal, Me
        If SchOrganization.ParaOutOrganizationID = "" Then FunSelectOrganizaton = False: Exit Function
       TxtOrganizationID.Text = SchOrganization.ParaOutOrganizationID
    End If
    If TxtOrganizationID.Text = "" Then FunSelectOrganizaton = False: Exit Function
    vStrSQL = " Select * FROM Organizations where OrganizationID='" & TxtOrganizationID.Text & "'"
    With CN.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtOrganizatonName.Text = !OrganizationName
          FunSelectOrganizaton = True
          .Close
          Exit Function
      Else
          FunSelectOrganizaton = False
          .Close
          TxtOrganizationID.Text = ""
          TxtOrganizatonName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function



