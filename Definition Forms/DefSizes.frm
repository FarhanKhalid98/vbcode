VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form DefSizes 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15450
   Icon            =   "DefSizes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   796
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1030
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
      Height          =   2850
      Left            =   12720
      TabIndex        =   14
      Top             =   3360
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
         Height          =   2445
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Tag             =   "NC"
         Text            =   "DefSizes.frx":0ECA
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
         TabIndex        =   16
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.TextBox TxtFilter 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2843
      MaxLength       =   30
      TabIndex        =   4
      Top             =   3083
      Width           =   4425
   End
   Begin VB.TextBox TxtID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8978
      MaxLength       =   5
      TabIndex        =   12
      Top             =   4883
      Width           =   1200
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4650
      Left            =   2843
      TabIndex        =   5
      Top             =   3413
      Width           =   4425
      ScrollBars      =   2
      _Version        =   196616
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "DefSizes.frx":0F55
      AllowUpdate     =   0   'False
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
      Columns.Count   =   2
      Columns(0).Width=   1852
      Columns(0).Caption=   "Size ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4895
      Columns(1).Caption=   "Size Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   7805
      _ExtentY        =   8202
      _StockProps     =   79
      BackColor       =   15724527
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8978
      MaxLength       =   50
      TabIndex        =   0
      Top             =   6023
      Width           =   3360
   End
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   2903
      TabIndex        =   6
      Top             =   8648
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "New"
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
      MICON           =   "DefSizes.frx":0F71
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   4223
      TabIndex        =   7
      Top             =   8648
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Change"
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
      MICON           =   "DefSizes.frx":0F8D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   5543
      TabIndex        =   8
      Top             =   8648
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Remove"
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
      MICON           =   "DefSizes.frx":0FA9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   8693
      TabIndex        =   1
      Top             =   8633
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
      MICON           =   "DefSizes.frx":0FC5
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   10013
      TabIndex        =   2
      Top             =   8633
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
      MICON           =   "DefSizes.frx":0FE1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   11333
      TabIndex        =   3
      Top             =   8633
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
      MICON           =   "DefSizes.frx":0FFD
      BC              =   14737632
      FC              =   0
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
      Left            =   10800
      TabIndex        =   17
      Top             =   630
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sizes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   13
      Top             =   270
      Width           =   705
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   530.533
      X2              =   529.533
      Y1              =   199.533
      Y2              =   541.533
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size Name"
      Height          =   195
      Left            =   2850
      TabIndex        =   11
      Top             =   2880
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size ID"
      Height          =   195
      Left            =   8985
      TabIndex        =   10
      Top             =   4665
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size Name"
      Height          =   195
      Left            =   8985
      TabIndex        =   9
      Top             =   5790
      Width           =   765
   End
End
Attribute VB_Name = "DefSizes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vMode As FormMode
Dim vMaxBinID As Integer
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.

Private Sub BtnClear_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('Sizes'" & "," & TxtID.Text & ",Null,'Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  FormStatus = SelectionMode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name = Grid.Name Then Call Grid_DblClick: Exit Sub
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      KeyCode = 0
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
          Case vbKeyH
             FraHelp.ZOrder 0
             FraHelp.Visible = True
             KeyCode = 0
          Case vbKeyN
              If BtnNew.Enabled Then BtnNew_Click
              KeyCode = 0
          Case vbKeyO
              If BtnOpen.Enabled Then BtnOpen_Click
              KeyCode = 0
          Case vbKeyR
              If BtnDelete.Enabled Then BtnDelete_Click
              KeyCode = 0
      End Select
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnSave.Enabled = False Then FormStatus = ChangeMode
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('Sizes'" & "," & TxtID.Text & ",Null,'Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniSizes", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
   Dim vtbl As String
   If Rs.RecordCount > 0 Then
    If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
      vtbl = Common.ChildDataExists("Sizes", "SizeID='" & Rs!SizeID & "'", "")
      If vtbl <> "" Then
         MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
         Exit Sub
      End If
      Call ActivityLog("Sizes", eDelete, , , TxtID.Text)
      
'    vMaxBinID = FunGetMaxBinID
'    ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'    CN.Execute ("Insert Into Bin_Sizes Select " & vMaxBinID & ",'" & Date & "',* from Sizes Where SizeID = " & TxtID.Text)
'
'    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    CN.Execute ("Insert Into UserActivities values ('Sizes'" & "," & TxtID.Text & ",Null,'Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      Rs.Delete
      If Rs.RecordCount = 0 Then FormStatus = NewMode: Exit Sub
      Rs.MoveNext
      Grid.MoveNext
      If Rs.EOF Then Rs.MoveLast
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnNew_Click()
   FormStatus = NewMode
End Sub

Private Sub BtnOpen_Click()
  On Error GoTo ErrorHandler
  If Rs.RecordCount > 0 Then
    If Rs.BOF = False And Rs.EOF = False Then
      FormStatus = OpenMode
    End If
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   If FunValidation = False Then Exit Sub
    
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniSizes", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If vIsNewRecord = False Then Call ActivityLog("Sizes", eEdit, , , TxtID.Text)
'   Call UserActivities
   Set Rs = New ADODB.Recordset
   Rs.Open " Select * FROM Sizes where SizeID = '" & TxtID.Text & "'", cn, adOpenDynamic, adLockOptimistic
   If vIsNewRecord Then
      Rs.AddNew
      Rs!SizeID = TxtID.Text
   End If
   Rs!SizeName = TxtName.Text
   Rs.Update
   If vIsNewRecord = True Then Call ActivityLog("Sizes", eAdd, , , TxtID.Text)
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
   On Error GoTo ErrorHandler
   If vIsNewRecord Then
      If Trim(TxtID.Text) = "" Then
         MsgBox "Please specify a Size ID", vbExclamation, "Alert"
         If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
         Exit Function
      End If
'      If Len(Trim(TxtID.Text)) < 3 Then
'         MsgBox "The Size ID must be three characters long", vbExclamation, "Alert"
'         TxtID.Text = Right("000" + CStr(Val(TxtID.Text)), 3)
'         If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
'         Exit Function
'      End If
      If cn.Execute("Select count(*) from Sizes where SizeID = '" & TxtID.Text & "'").Fields(0) > 0 Then
          MsgBox "This Size ID already exists. The Size ID must be unique", vbExclamation, "Alert"
          TxtID.Text = FunGetMaxID
          If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
          Exit Function
      End If
      Select Case Asc(UCase(Left(TxtID.Text, 1)))
        Case 65 To 90
        Case 48 To 57
        Case Else
          MsgBox "The Size ID must contain numeric/alphabetical characters only", vbExclamation, "Alert"
          If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
          Exit Function
      End Select
      Select Case Asc(UCase(Right(TxtID.Text, 1)))
        Case 65 To 90
        Case 48 To 57
        Case Else
          MsgBox "The Size ID must contain numeric/alphabetical characters only", vbExclamation, "Alert"
          If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
          Exit Function
      End Select
      If cn.Execute("Select * FROM Sizes where SizeName like '" & TxtName.Text & "'").RecordCount > 0 Then
         MsgBox "This Size Already Exists", vbExclamation, "Alert"
         If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
         Exit Function
      End If
   Else
      If cn.Execute("Select * FROM Sizes where SizeID <> '" & TxtID.Text & "' and SizeName like '" & TxtName.Text & "'").RecordCount > 0 Then
         MsgBox "This Size Already Exists", vbExclamation, "Alert"
         If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
         Exit Function
      End If
   End If
   If Trim(TxtName.Text) = "" Then
      MsgBox "Please specify a Size name", vbExclamation, "Alert"
      If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
      Exit Function
   End If
   'All Ok, now validation is success
   FunValidation = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub Grid_GotFocus()
   On Error GoTo ErrorHandler
   Call Grid_RowColChange(0, 0)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LblClose_Click()
   FraHelp.Visible = False
End Sub

Private Sub LblHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   LblHelp.ForeColor = &H800000
   FraHelp.ZOrder 0
   FraHelp.Visible = True
End Sub

Private Sub LblHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If LblHelp.FontUnderline = True Then Exit Sub
   LblHelp.FontUnderline = True
End Sub

Private Sub LblHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   LblHelp.ForeColor = vbWhite
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
   SetWindowText Me.hWnd, "Sizes"
   HelpLocation Me
   Set Rs = New ADODB.Recordset
   Rs.Open "Select * FROM Sizes order by SizeName", cn, adOpenDynamic, adLockOptimistic
   Set Grid.DataSource = Rs
   Grid.Columns("ID").DataField = "SizeID"
   Grid.Columns("Name").DataField = "SizeName"
   FormStatus = NewMode
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   Exit Sub
ErrorHandler:
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
         BtnNew.Enabled = False
         BtnOpen.Enabled = False
         BtnDelete.Enabled = False
         BtnSave.Enabled = False
         BtnClear.Enabled = True
         TxtName.Enabled = False
         TxtID.Enabled = False
         TxtName.Text = ""
         TxtID.Text = ""
         TxtFilter.Text = ""
         TxtFilter.Enabled = False
         TxtFilter.BackColor = &HE0E0E0
         TxtID.Text = FunGetMaxID()
         TxtName.Enabled = True
         TxtName.BackColor = vbWhite
         TxtID.Enabled = True
         TxtID.BackColor = vbWhite
         Grid.Enabled = False
         'If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
         If TxtName.Visible = False Then TxtName.Visible = True
         If TxtName.Visible And TxtName.Enabled Then TxtName.SetFocus
         vIsNewRecord = True
     Case Is = OpenMode
         BtnNew.Enabled = False
         BtnOpen.Enabled = False
         BtnDelete.Enabled = False
         BtnClear.Enabled = True
         Grid.Enabled = False
         Call Grid_RowColChange(0, 0)
         TxtFilter.Enabled = False
         TxtFilter.BackColor = &HE0E0E0
         TxtName.Enabled = True
         TxtID.Enabled = False
         TxtID.BackColor = &HE0E0E0
         TxtName.BackColor = vbWhite
         TxtName.SetFocus
         TxtFilter.Text = ""
         vIsNewRecord = False
     Case Is = ChangeMode
         BtnSave.Enabled = True
     Case Is = SelectionMode
         Grid.Enabled = True
         'TxtFilter.Text = ""
         TxtFilter.Enabled = True
         TxtFilter.BackColor = vbWhite
         BtnNew.Enabled = True
         BtnOpen.Enabled = True
         BtnDelete.Enabled = True
         BtnSave.Enabled = False
         BtnClear.Enabled = False
         TxtName.Enabled = False
         TxtName.BackColor = &HE0E0E0
         TxtID.Enabled = False
         TxtID.BackColor = &HE0E0E0
         Call Grid_RowColChange(0, 0)
         Grid.SetFocus
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   Else
      Dim frmObj As Object
      For Each frmObj In Forms
          Set frmObj = Nothing
      Next
      Set Rs = Nothing
      Set DefSizes = Nothing
   End If
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKeyA To vbKeyZ, vbKey0 To vbKey9, Asc("a") To Asc("z")
      TxtFilter.Text = Chr(KeyAscii): TxtFilter.SelStart = Len(TxtFilter.Text):  TxtFilter.SetFocus
   End Select
   
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo ErrorHandler
   If Rs.RecordCount > 0 And Grid.Enabled Then
      TxtID.Text = Grid.Columns("ID").Text
      TxtName.Text = Grid.Columns("Name").Text
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
   If Grid.Rows > 0 And BtnOpen.Enabled Then BtnOpen_Click
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtFilter_Change()
   On Error GoTo ErrorHandler
   'If Me.ActiveControl.Name <> TxtFilter.Name Then Exit Sub
   'If Trim(TxtFilter.Text) = "" Then Grid.MoveFirst: Exit Sub
   'Rs.Find "SizeName like '" & Replace(TxtFilter.Text, "'", "''") & "%'", , adSearchForward, 1
   Set Rs = New ADODB.Recordset
   Rs.Open " Select * FROM Sizes where SizeName like '%" & Replace(TxtFilter.Text, "'", "''") & "%' Order by SizeName", cn, adOpenStatic, adLockOptimistic
   Set Grid.DataSource = Rs
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As String
  On Error GoTo ErrorHandler
  FunGetMaxID = cn.Execute("Select isnull(max(SizeID),0) + 1  from Sizes").Fields(0)
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub TxtName_Change()
   If Me.ActiveControl.Name <> TxtName.Name Then Exit Sub
   TxtFilter.Text = TxtName.Text
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxBinID = cn.Execute("Select isnull(max(BinID),0)+1 from Bin_Sizes ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
    If vIsNewRecord = False Then
    With cn.Execute("Select  * from Sizes where SizeID =" & TxtID.Text)
        If TxtName.Text <> !SizeName Then
            cn.Execute ("Insert Into UserActivities values ('Sizes'" & "," & TxtID.Text & ", Null , 'Updated SizeName-" & !SizeName & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        
    End With
   Else
        cn.Execute ("Insert Into UserActivities values ('Sizes'" & "," & TxtID.Text & ", Null ,'Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

