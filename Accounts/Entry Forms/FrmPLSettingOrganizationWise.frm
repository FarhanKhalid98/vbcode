VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmPLSettingsOrganizationWise 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmPLSettingOrganizationWise.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbOrganization 
      Height          =   315
      Left            =   7508
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2400
      Width           =   2145
   End
   Begin VB.Frame FraHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   12120
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   4200
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Tag             =   "NC"
         Text            =   "FrmPLSettingOrganizationWise.frx":0ECA
         Top             =   360
         Width           =   3930
      End
      Begin VB.Label LblClose 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3915
         TabIndex        =   10
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.TextBox TxtAccountName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3518
      TabIndex        =   1
      Top             =   2430
      Width           =   3255
   End
   Begin VB.ComboBox CmbFilter 
      Height          =   315
      Left            =   9683
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2400
      Width           =   2160
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6015
      TabIndex        =   2
      Top             =   8535
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save"
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
      MICON           =   "FrmPLSettingOrganizationWise.frx":0F1C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7320
      TabIndex        =   3
      Top             =   8535
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Clear"
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
      MICON           =   "FrmPLSettingOrganizationWise.frx":0F38
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8655
      TabIndex        =   4
      Top             =   8535
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
      MICON           =   "FrmPLSettingOrganizationWise.frx":0F54
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5280
      Left            =   3525
      TabIndex        =   14
      Top             =   2760
      Width           =   8310
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
      stylesets.count =   2
      stylesets(0).Name=   "SelectedCol"
      stylesets(0).BackColor=   12713983
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
      stylesets(0).Picture=   "FrmPLSettingOrganizationWise.frx":0F70
      stylesets(1).Name=   "SelectedRow"
      stylesets(1).ForeColor=   16777215
      stylesets(1).BackColor=   8388608
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "FrmPLSettingOrganizationWise.frx":0F8C
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedCol"
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      Columns.Count   =   6
      Columns(0).Width=   1852
      Columns(0).Caption=   "A/c No."
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   8493
      Columns(1).Caption=   "A/c Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3200
      Columns(2).Caption=   "Setting"
      Columns(2).Name =   "Setting"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(2).Style=   3
      Columns(2).Row.Count=   3
      Columns(2).Col.Count=   2
      Columns(2).Row(0).Col(0)=   "+"
      Columns(2).Row(1).Col(0)=   "-"
      Columns(2).Row(2).Col(0)=   " "
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "Narration"
      Columns(3).Name =   "Narration"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "Debit"
      Columns(4).Name =   "Debit"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "Credit"
      Columns(5).Name =   "Credit"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   14658
      _ExtentY        =   9313
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7530
      TabIndex        =   13
      Top             =   2175
      Width           =   1080
   End
   Begin VB.Label LblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   11070
      TabIndex        =   11
      Top             =   495
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P/L Setting"
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
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3585
      TabIndex        =   6
      Top             =   2205
      Width           =   1065
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9705
      TabIndex        =   5
      Top             =   2175
      Width           =   1080
   End
End
Attribute VB_Name = "FrmPLSettingsOrganizationWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vSuppressUpdateEvent As Boolean
Dim vSQL As String

Private Sub cmbfilter_click()
  On Error GoTo ErrorHandler
  If Rs.State = adStateOpen Then
    Rs.CancelBatch
    Rs.Close
  End If
  Me.MousePointer = vbHourglass
  Rs.Open "Select * FROM OrganizationChartOfAccounts Where OrganizationID = " & CmbOrganization.ItemData(CmbOrganization.ListIndex), cn, adOpenStatic, adLockBatchOptimistic
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
  vSQL = "Select Isnull(org.accountNo,c.accountNo) AccountNo, AccountName, " & vbCrLf _
       + " Isnull(org.Narration,c.Narration) Narration, " & vbCrLf _
       + " Isnull(org.openingdebit,c.openingdebit) OpeningDebit," & vbCrLf _
       + " Isnull(org.Openingcredit,c.Openingcredit) Openingcredit, " & vbCrLf _
       + " Isnull(Isnull(org.PlFlag,c.PlFlag),'') as PlFlag, " & vbCrLf _
       + " org.organizationID From" & vbCrLf _
       + " (" & vbCrLf _
       + " Select o.* FROM ChartOfAccounts C" & vbCrLf _
       + " Left Outer Join OrganizationChartOfAccounts O" & vbCrLf _
       + " On O.AccountNo = C.AccountNo Where  OrganizationID = " & CmbOrganization.ItemData(CmbOrganization.ListIndex) & vbCrLf _
       + " )Org" & vbCrLf _
       + " right Outer Join ChartOfAccounts C on C.AccountNo = Org.AccountNo" & vbCrLf _
       + " Where IsDetailed=1 " & IIf(CmbFilter.ListIndex = 0, "", " And AccountType = '" & CmbFilter.Text & "'") & " and AccountName like '%" & TxtAccountName.Text & "%'"
    With cn.Execute(vSQL)
    Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !AccountNo
        Grid.Columns("Name").Text = !AccountName
        Grid.Columns("Narration").Text = IIf(IsNull(!Narration), "", !Narration)
        Grid.Columns("Debit").Value = !openingdebit
        Grid.Columns("Credit").Value = !openingCredit
        Grid.Columns("Setting").Text = !plflag
        Grid.Update
        .MoveNext
    Loop
  End With
  vSuppressUpdateEvent = False
  Grid.Redraw = True
  Grid.MoveFirst
  'If Grid.Visible Then Grid.SetFocus
  Me.MousePointer = vbDefault
  Exit Sub
ErrorHandler:
  Grid.Redraw = True
  Me.MousePointer = vbDefault
  Call ShowErrorMessage
End Sub

Private Sub BtnClear_Click()
  Call cmbfilter_click
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   Grid.Update
   Rs.Filter = 0
   Rs.MoveFirst
   While Not Rs.EOF
      If Rs.EditMode <> adEditNone Then
         Call ActivityLog("Account Opening Balance", eEdit, , , Rs!AccountNo)
      End If
      Rs.MoveNext
   Wend
   Rs.UpdateBatch
   MsgBox "Your Entries has been Successfully Updated.", vbOKOnly + vbInformation, "Information"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbOrganization_Change()
 If CmbOrganization.Visible = False Then Exit Sub
 If Me.ActiveControl.Name <> CmbOrganization.Name Then Exit Sub
 Call cmbfilter_click
End Sub

Private Sub CmbOrganization_Click()
 If CmbOrganization.Visible = False Then Exit Sub
 If Me.ActiveControl.Name <> CmbOrganization.Name Then Exit Sub
 Call cmbfilter_click
End Sub

Private Sub Grid_LostFocus()
   On Error GoTo ErrorHandler
   If vSuppressUpdateEvent Then Exit Sub
   If Grid.Visible = False Then Exit Sub
   UpdateOrganizationAccounts
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LblClose_Click()
   FraHelp.Visible = False
End Sub

Private Sub LblHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LblHelp.ForeColor = &H800000
   FraHelp.ZOrder 0
   FraHelp.Visible = True
End Sub

Private Sub LblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If LblHelp.FontUnderline = True Then Exit Sub
   LblHelp.FontUnderline = True
End Sub

Private Sub LblHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LblHelp.ForeColor = vbWhite
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
   If LblHelp.FontUnderline = False Then Exit Sub
   LblHelp.FontUnderline = False
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
  ShowPicture Me, 2
  
  AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
  SetWindowText Me.hWnd, "Accounts Opening Balance"
  HelpLocation Me
  
   ''''''''''''''''''''''''''' Organizations '''''''''''''''''''''''
   With cn.Execute("Select * from Organizations")
      While Not .EOF
         CmbOrganization.AddItem !OrganizationName
         CmbOrganization.ItemData(CmbOrganization.NewIndex) = !OrganizationID
         .MoveNext
      Wend
      .Close
   End With
   If CmbOrganization.ListCount > 0 Then CmbOrganization.ListIndex = 0
   
  ''''''''''''''''''''''''''' Account Type '''''''''''''''''''''''
  CmbFilter.AddItem "All Accounts"
  With cn.Execute("Select Distinct AccountType from ChartofAccounts Where isdetailed = 1")
    Do Until .EOF
      CmbFilter.AddItem !AccountType
      .MoveNext
    Loop
  End With
  'CmbFilter.AddItem "Party"
  If CmbFilter.ListCount > 0 Then CmbFilter.ListIndex = 0
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name <> Grid.Name Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnSave.Enabled Then BtnSave_Click
            KeyCode = 0
         Case vbKeyH
               FraHelp.ZOrder 0
               FraHelp.Visible = True
               KeyCode = 0
         Case vbKeyW
            If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   End If
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  'If Grid.Columns(ColIndex).Text = "" Then Grid.Columns(ColIndex).Text = "0"
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
  On Error GoTo ErrorHandler
  If vSuppressUpdateEvent Then Exit Sub
    If Grid.Visible = False Then Exit Sub
   If ActiveControl.Name <> Grid.Name Then Exit Sub
   UpdateOrganizationAccounts
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub UpdateOrganizationAccounts()
'  If Val(Grid.Columns("Multiplier").Value) = 0 Then Grid.Columns("Multiplier").Value = 0
   Rs.Filter = "AccountNo = " & Val(Grid.Columns("ID").Value)
   If Rs.RecordCount = 0 Then
      Rs.AddNew
      Rs!OrganizationID = CmbOrganization.ItemData(CmbOrganization.ListIndex)
      Rs!AccountNo = Grid.Columns("ID").Value
      Rs!Narration = Grid.Columns("Narration").Value
      Rs!openingdebit = Val(Grid.Columns("Debit").Value)
      Rs!openingCredit = Val(Grid.Columns("Credit").Value)
      Rs!plflag = Grid.Columns("setting").Text
      Rs!UserNo = vUser
'      If vIsNewRecord = False Then CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Inserted New AccountNo-v" & Rs!AccountNo & " Multiplier- " & Rs!Multiplier & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
   ElseIf Rs.RecordCount = 1 Then
      Rs!Narration = Val(Grid.Columns("Narration").Value)
      Rs!openingdebit = Val(Grid.Columns("Debit").Value)
      Rs!openingCredit = Val(Grid.Columns("Credit").Value)
      Rs!plflag = Grid.Columns("setting").Text
      Rs!UserNo = vUser
'      If vIsNewRecord = False Then CN.Execute ("Insert Into UserActivities values ('Products'" & "," & TxtID.Text & ", Null , 'Updated AccountNo-v" & Rs!AccountNo & " Multiplier- " & Rs!Multiplier & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
      Rs.Update
  End If
End Sub

Private Sub Grid_GotFocus()
   Grid.Row = 0
   Grid.Col = 0
   SendKeys "{Right}"
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event vbKeyRight, 1, 1, 1
      KeyCode = 0
   End If
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtAccountName_Change()
   On Error GoTo ErrorHandler
   Call cmbfilter_click
   Exit Sub
ErrorHandler:
  Grid.Redraw = True
  'Me.MousePointer = vbDefault
  Call ShowErrorMessage
End Sub

