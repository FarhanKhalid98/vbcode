VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form ExpSettings 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "ExpSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
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
      Left            =   11880
      TabIndex        =   6
      Top             =   1200
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
         TabIndex        =   7
         Tag             =   "NC"
         Text            =   "ExpSettings.frx":0ECA
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
         TabIndex        =   8
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.ComboBox CmbFilter 
      Height          =   315
      Left            =   3968
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2250
      Width           =   2160
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8385
      TabIndex        =   0
      Top             =   8685
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
      MICON           =   "ExpSettings.frx":0F1C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   5700
      TabIndex        =   3
      Top             =   8685
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
      MICON           =   "ExpSettings.frx":0F38
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7042
      TabIndex        =   4
      Top             =   8685
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
      MICON           =   "ExpSettings.frx":0F54
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5640
      Left            =   3968
      TabIndex        =   10
      Top             =   2670
      Width           =   7425
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   4
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
      stylesets(0).Picture=   "ExpSettings.frx":0F70
      CheckBox3D      =   0   'False
      MultiLine       =   0   'False
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
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   4
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "A/c No."
      Columns(0).Name =   "TaskKey"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1773
      Columns(1).Caption=   "A/c No."
      Columns(1).Name =   "ID"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   7064
      Columns(2).Caption=   "Account Name"
      Columns(2).Name =   "Name"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   3200
      Columns(3).Caption=   "Allow"
      Columns(3).Name =   "Setting"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Style=   2
      _ExtentX        =   13097
      _ExtentY        =   9948
      _StockProps     =   79
      Caption         =   "Expense Setting"
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
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11610
      Top             =   45
      Width           =   330
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
      Left            =   10935
      TabIndex        =   9
      Top             =   855
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Setting"
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
      TabIndex        =   5
      Top             =   270
      Width           =   2310
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type:"
      Height          =   225
      Left            =   3968
      TabIndex        =   2
      Top             =   2025
      Width           =   1170
   End
   Begin VB.Menu MnuAccounts 
      Caption         =   "Accounts"
      Visible         =   0   'False
      Begin VB.Menu mniCreateNewChildGroup 
         Caption         =   "Create new Child Group"
      End
      Begin VB.Menu mniModifyPropertiesForGroup 
         Caption         =   "Modify the Properties for this Group"
      End
      Begin VB.Menu mniDeleteGroup 
         Caption         =   "Delete this Group"
      End
      Begin VB.Menu MniSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mniCreatenewAccounts 
         Caption         =   "Create new Accounts for this Group"
      End
      Begin VB.Menu mniModifyAccount 
         Caption         =   "Modify the Selected Account"
      End
      Begin VB.Menu mniDeleteAccount 
         Caption         =   "Delete this Account"
      End
   End
End
Attribute VB_Name = "ExpSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vSuppressUpdateEvent As Boolean
Dim vSQL As String
'----------------------------------

Private Sub cmbfilter_click()
  On Error GoTo ErrorHandler
  If Rs.State = adStateOpen Then
    Rs.CancelBatch
    Rs.Close
  End If
  Me.MousePointer = vbHourglass
  vSQL = " Select c.AccountNo, AccountName, case when e.AccountNo is null then 0 else 1 end as ExpFlag" & vbCrLf _
      + " from ChartofAccounts c left outer join ExpenseAccounts e on c.AccountNo = e.AccountNo" & vbCrLf _
      + " where IsDetailed=1 and AccountType = '" & CmbFilter.Text & "'"
      
  Rs.Open "Select * from ExpenseAccounts", cn, adOpenDynamic, adLockBatchOptimistic
  
  Grid.Redraw = False
  Grid.CancelUpdate
  Grid.RemoveAll
  vSuppressUpdateEvent = True
  With cn.Execute(vSQL)
    Do Until .EOF
        Grid.AddNew
        Grid.Columns("ID").Text = !AccountNo
        Grid.Columns("Name").Text = !AccountName
        Grid.Columns("Setting").Text = !ExpFlag
        Grid.Update
        .MoveNext
    Loop
  End With
  vSuppressUpdateEvent = False
  Grid.MoveFirst
  Grid.Redraw = True
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
  Rs.Filter = ""
'  Rs.MoveFirst
'  While Not Rs.EOF
'      If Rs.EditMode <> adEditNone Then
'         Call ActivityLog("Exp Setting", eEdit, , , Rs!AccountNo)
'      End If
'      Rs.MoveNext
'  Wend
  Rs.UpdateBatch
  MsgBox "Your Entries has been Successfully Updated.", vbOKOnly + vbInformation, "Information"
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
'    MsgBox Grid.Columns("Setting").Value
End Sub

Private Sub ImgExit_Click()
   Unload Me
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
  SetWindowText Me.hWnd, "Exp Setting"
  HelpLocation Me
  With cn.Execute("Select Distinct AccountType from ChartofAccounts Where AccountDepth = 0")
    Do Until .EOF
      CmbFilter.AddItem !AccountType
      .MoveNext
    Loop
  End With
  If CmbFilter.ListCount > 0 Then CmbFilter.Text = "Expenses"
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnSave.Enabled Then BtnSave_Click
            KeyCode = 0
         Case vbKeyW
            If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
         Case vbKeyH
               FraHelp.ZOrder 0
               FraHelp.Visible = True
               KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   End If
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
   If Grid.Columns(ColIndex).Text = "" Then Grid.Columns(ColIndex).Text = "0"
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   On Error GoTo ErrorHandler
   If vSuppressUpdateEvent Then Exit Sub
   UpdateExpense
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub UpdateExpense()
   If Val(Grid.Columns("Setting").Value) = 0 Then Grid.Columns("Setting").Value = 0
   Rs.Filter = "AccountNo = " & Val(Grid.Columns("ID").Text)
   If Rs.RecordCount = 0 Then
      Rs.AddNew
      Rs!AccountNo = Grid.Columns("ID").Text
      Rs.Update
'            CN.Execute ("Insert Into ExpenseAccounts values '" & Grid.Columns("ID").Text & "'")
   ElseIf Rs.RecordCount = 1 And Val(Grid.Columns("Setting").Value) = 0 Then
'     If vIsNewRecord = False Then CN.Execute ("Insert Into ExpenseAccounts values ('Products'" & "," & TxtID.Text & ", Null , 'Deleted PackingID-v" & Rs!PackingID & " Multiplier- " & Rs!Multiplier & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
      Rs.Delete
      Rs.Update
  End If
End Sub

