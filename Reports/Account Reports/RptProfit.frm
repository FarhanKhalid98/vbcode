VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RptProfit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "RptProfit.frx":0000
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
      Left            =   8446
      TabIndex        =   6
      Top             =   7348
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
      MICON           =   "RptProfit.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5626
      TabIndex        =   4
      Top             =   7348
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
      MICON           =   "RptProfit.frx":0EE6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7006
      TabIndex        =   5
      Top             =   7348
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
      MICON           =   "RptProfit.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnGroup 
      Height          =   330
      Left            =   5993
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3588
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
      MICON           =   "RptProfit.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtGroupID 
      Height          =   315
      Left            =   4658
      TabIndex        =   0
      Top             =   3588
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   3
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
   Begin SITextBox.Txt TxtGroupName 
      Height          =   315
      Left            =   6353
      TabIndex        =   8
      Top             =   3588
      Width           =   4350
      _ExtentX        =   7673
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
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   4658
      TabIndex        =   1
      Top             =   4383
      Width           =   1335
      _ExtentX        =   2355
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
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   5993
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4368
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
      MICON           =   "RptProfit.frx":0F3A
      BC              =   12632256
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   6353
      TabIndex        =   12
      Top             =   4383
      Width           =   4350
      _ExtentX        =   7673
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
      Masked          =   5
   End
   Begin MSComCtl2.DTPicker DtpFrom 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   5498
      TabIndex        =   2
      Top             =   5673
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   130744323
      CurrentDate     =   38244
   End
   Begin MSComCtl2.DTPicker DtpTo 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   7268
      TabIndex        =   3
      Top             =   5673
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   130744323
      CurrentDate     =   38244
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5498
      TabIndex        =   17
      Top             =   5448
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7268
      TabIndex        =   16
      Top             =   5448
      Width           =   1095
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profit Group wise"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   15
      Top             =   270
      Width           =   3030
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6353
      TabIndex        =   14
      Top             =   4143
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4658
      TabIndex        =   13
      Top             =   4143
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6353
      TabIndex        =   10
      Top             =   3378
      Width           =   1065
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4658
      TabIndex        =   9
      Top             =   3363
      Width           =   780
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
Attribute VB_Name = "RptProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String, vDate As String
Dim RsReport As New ADODB.Recordset
Dim vStockLimit As String
Dim vGroup As String
Dim vProduct As String
Dim VStrSQL As String

Private Sub BtnGroup_Click()
   If FunSelectGroup(ssButton, False) = True Then
      TxtProductID.SetFocus
   Else
      TxtGroupID.SetFocus
   End If
End Sub

Private Sub BtnClose_Click()
    Unload Me
End Sub

Private Sub BtnPreview_Click()
    If SetReport Then
        RptReportViewer.Caption = Me.Caption
        RptReportViewer.Show vbModal
    End If
End Sub

Private Sub BtnPrint_Click()
    If SetReport Then RptReportViewer.Report.PrintOut False
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, False) = True Then
      BtnPreview.SetFocus
   Else
      TxtProductID.SetFocus
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
           Case TxtGroupID.Name: If FunSelectGroup(ssFunctionKey, True) = True Then TxtProductID.SetFocus
           Case TxtProductID.Name: If FunSelectProduct(ssFunctionKey, True) = True Then BtnPreview.SetFocus
      End Select
    End If
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Profit Product Wise Report"
   DtpFrom.Value = Date - 30
   DtpTo.Value = Date
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
   Set RsReport = Nothing
   Set RptProfit = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function SetReport() As Boolean
   On Error GoTo ErrorHandler
   SetReport = False
   vGroup = IIf(TxtGroupID.Text = "", "", " and p.groupid='" & TxtGroupID.Text & "'")
   vProduct = IIf(TxtProductID.Text = "", "", " and d.Productid='" & TxtProductID.Text & "'")
   sSql = " select d.Productid, ProductName, isnull(dbo.FunPurPrice(d.Productid),p.purprice) as PurPrice, RetailPrice, p.GroupiD,GroupName, sum(samount) - sum(ramount) as SAmount, sum(sqty) - sum(RQty) as Qty from " & _
         " (select b.productid, sum(amount) as SAmount, 0 as RAmount ,sum(qty)sQty, 0 as RQTy from salebody b inner join Products p on b.productid=p.productid" & _
         " where Billdate between '" & DtpFrom.Value & "' and '" & DtpTo.Value & "'" & _
         " group by b.productid" & _
         " Union All" & _
         " select b.productid, 0, sum(Amount) as RAmount, 0, sum(qty) from salereturnbody b  inner join Products p on b.productid=p.productid" & _
         " where ReturnDate between '" & DtpFrom.Value & "' and '" & DtpTo.Value & "'" & _
         " group by b.productid) d inner join products p on p.productid = d.productid inner join groups g on g.groupid = p.groupid" & _
         " where 1=1" & vGroup & vProduct & _
         " group by d.Productid, ProductName, PurPrice, RetailPrice, p.GroupID, GroupName"
          
   Me.MousePointer = vbHourglass
   If RsReport.State = adStateOpen Then RsReport.Close
   RsReport.Open sSql, CN, adOpenStatic, adLockReadOnly
   If RsReport.BOF Then
       MsgBox "No record exists.", vbInformation, Me.Caption
       Me.MousePointer = vbDefault
       Exit Function
   End If
   Set RptReportViewer.Report = New CrpProfit
   RptReportViewer.Report.Database.SetDataSource RsReport
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyName
   RptReportViewer.Report.ParameterFields(2).AddCurrentValue ObjRegistry.CompanyAddress & IIf(IsNull(ObjRegistry.CompanyCity), "", ", " & ObjRegistry.CompanyCity)
   RptReportViewer.Report.ParameterFields(3).AddCurrentValue IIf(ObjRegistry.CompanyPhoneNo = "", "", "Phone # " & ObjRegistry.CompanyPhoneNo)
   RptReportViewer.Report.ParameterFields(4).AddCurrentValue ObjRegistry.DevelopedBy
   RptReportViewer.Report.ParameterFields(5).AddCurrentValue "From : " & Format(DtpFrom.Value, "dd/MM/yyyy") & ",   To : " & Format(DtpTo.Value, "dd/MM/yyyy")
   RptReportViewer.Report.SelectPrinter ObjRegistry.DriverName, ObjRegistry.DeviceName, ObjRegistry.Port
   SetReport = True
   Me.MousePointer = vbDefault
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtGroupID_Change()
   If TxtGroupID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtGroupID.Name Then Exit Sub
   If TxtGroupName.Text <> "All Groups" Then
      TxtGroupName.Text = "All Groups"
   End If
End Sub

Private Sub TxtGroupID_Validate(Cancel As Boolean)
If Me.ActiveControl.Name <> TxtGroupID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   vGroup = ""
   If TxtGroupName.Text <> "All Groups" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectGroup(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectGroup(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectGroup(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    '-- when Group ID is written then it will check and all its related value will be write its appropriate places
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchGroup.Show vbModal, Me
        If SchGroup.ParaOutGroupID = "" Then FunSelectGroup = False: Exit Function
        TxtGroupID.Text = SchGroup.ParaOutGroupID
    End If
    '---------------------------
    VStrSQL = " Select * FROM Groups where GroupID='" & TxtGroupID.Text & "'"
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtGroupName.Text = !GroupName
          FunSelectGroup = True
          .Close
          Exit Function
      Else
          FunSelectGroup = False
          .Close
          TxtGroupID.Text = ""
          TxtGroupName.Text = ""
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtProductID_Change()
   If TxtProductID.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtProductID.Name Then Exit Sub
   If TxtProductName.Text <> "All Products" Then
      TxtProductName.Text = "All Products"
   End If
End Sub

Private Sub TxtProductID_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtProductID.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtProductName.Text <> "All Products" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectProduct(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectProduct(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    '-- when Party ID is written then it will check and all its related value will be write its appropriate places
    Dim VStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchProduct.Show vbModal, Me
        If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
        TxtProductID.Text = SchProduct.ParaOutID
    End If
    '---------------------------
    VStrSQL = " Select * FROM Products where ProductID = '" & TxtProductID.Text & "'"
    With CN.Execute(VStrSQL)
      If .RecordCount > 0 Then
          TxtProductName.Text = !ProductName
          FunSelectProduct = True
          .Close
          Exit Function
      Else
          FunSelectProduct = False
          .Close
          TxtProductID.Text = ""
          TxtProductName.Text = "All Products"
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function
