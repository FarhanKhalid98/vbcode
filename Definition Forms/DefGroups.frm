VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form DefGroups 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15450
   Icon            =   "DefGroups.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   770
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkIsKitchen 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Is Kitchen"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9000
      TabIndex        =   23
      Top             =   8010
      Width           =   2850
   End
   Begin VB.TextBox TxtPCTCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   10755
      MaxLength       =   8
      TabIndex        =   21
      Top             =   4680
      Width           =   1200
   End
   Begin VB.CheckBox ChkisRemarksCompulsory 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Remarks compulsory in Sale"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9000
      TabIndex        =   20
      Top             =   7605
      Width           =   2850
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
      Height          =   2850
      Left            =   10845
      TabIndex        =   15
      Top             =   945
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
         TabIndex        =   16
         Tag             =   "NC"
         Text            =   "DefGroups.frx":0ECA
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
         TabIndex        =   17
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.TextBox TxtFilter 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2730
      MaxLength       =   30
      TabIndex        =   5
      Top             =   3135
      Width           =   4425
   End
   Begin VB.TextBox TxtID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8978
      MaxLength       =   4
      TabIndex        =   13
      Top             =   4688
      Width           =   1200
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4650
      Left            =   2730
      TabIndex        =   6
      Top             =   3465
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
      stylesets(0).Picture=   "DefGroups.frx":0F55
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
      Columns(0).Caption=   "Group ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4895
      Columns(1).Caption=   "Group Name"
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
      Top             =   5828
      Width           =   3360
   End
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   2903
      TabIndex        =   7
      Top             =   8453
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
      MICON           =   "DefGroups.frx":0F71
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   4223
      TabIndex        =   8
      Top             =   8453
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
      MICON           =   "DefGroups.frx":0F8D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   5543
      TabIndex        =   9
      Top             =   8453
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
      MICON           =   "DefGroups.frx":0FA9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   8693
      TabIndex        =   2
      Top             =   8438
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
      MICON           =   "DefGroups.frx":0FC5
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   10013
      TabIndex        =   3
      Top             =   8438
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
      MICON           =   "DefGroups.frx":0FE1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   11333
      TabIndex        =   4
      Top             =   8438
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
      MICON           =   "DefGroups.frx":0FFD
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PCT Code"
      Height          =   195
      Left            =   10755
      TabIndex        =   22
      Top             =   4455
      Width           =   735
   End
   Begin VB.Label LblGroupName1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name Urdu"
      Height          =   195
      Left            =   9000
      TabIndex        =   19
      Top             =   6600
      Width           =   1290
   End
   Begin MSForms.TextBox TxtGroupName1 
      Height          =   435
      Left            =   8985
      TabIndex        =   1
      ToolTipText     =   "Textbox1"
      Top             =   6840
      Width           =   4785
      VariousPropertyBits=   752896027
      ForeColor       =   0
      BorderStyle     =   1
      Size            =   "8440;767"
      SpecialEffect   =   0
      FontName        =   "@Arial Unicode MS"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      TabIndex        =   18
      Top             =   630
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Groups"
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
      Height          =   360
      Index           =   0
      Left            =   2700
      TabIndex        =   14
      Top             =   270
      Width           =   1035
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
      Y1              =   186.533
      Y2              =   528.533
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
      Height          =   195
      Left            =   2730
      TabIndex        =   12
      Top             =   2925
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group ID"
      Height          =   195
      Left            =   8978
      TabIndex        =   11
      Top             =   4463
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
      Height          =   195
      Left            =   8978
      TabIndex        =   10
      Top             =   5588
      Width           =   900
   End
End
Attribute VB_Name = "DefGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vMode As FormMode
Dim vMaxBinID As Integer
Dim ModeValue As Boolean
Dim UniCode As Variant
Dim vStrSQL As String
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.

Private Sub BtnClear_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('Groups'" & "," & TxtID.Text & ",Null,'Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  FormStatus = SelectionMode
End Sub

Private Sub ChkIsKitchen_Click()
On Error GoTo ErrorHandler
   If ActiveControl.Name <> ChkIsKitchen.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ChkisRemarksCompulsory_Click()
On Error GoTo ErrorHandler
   If ActiveControl.Name <> ChkisRemarksCompulsory.Name Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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
    cn.Execute ("Insert Into UserActivities values ('Groups'" & "," & TxtID.Text & ",Null,'Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
   
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniGroups", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
   Dim vtbl As String
   If Rs.RecordCount > 0 Then
    If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
      vtbl = Common.ChildDataExists("Groups", "GroupId='" & Rs!GroupID & "'", "")
      If vtbl <> "" Then
         MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
         Exit Sub
      End If
      Call ActivityLog("Groups", eDelete, , , TxtID.Text)
      
'    vMaxBinID = FunGetMaxBinID
'    ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header-----------------------------------------------
'    CN.Execute ("Insert Into Bin_Groups Select " & vMaxBinID & ",'" & Date & "',* from Groups Where GroupID = " & TxtID.Text)
'
'    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    CN.Execute ("Insert Into UserActivities values ('Groups'" & "," & TxtID.Text & ",Null,'Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
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
   vUserAction = UserAuthentication("MniGroups", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If vIsNewRecord = False Then Call ActivityLog("Groups", eEdit, , , TxtID.Text)
'   Call UserActivities
'   Set Rs = New ADODB.Recordset
'   Rs.Open " Select * FROM Groups where GroupID = '" & TxtID.Text & "'", cn, adOpenDynamic, adLockOptimistic
   Rs.Filter = "GroupID = '" & TxtID.Text & "'"
   If vIsNewRecord Then
      Rs.AddNew
      Rs!GroupID = TxtID.Text
'      Rs!isChanged = 0
   Else
      Rs!IsSync = 0
'      Rs!isChanged = 1
   End If
   Rs!GroupName = TxtName.Text
   Rs!GroupName1 = IIf(Trim(TxtGroupName1.Text) = "", Null, TxtGroupName1.Text)
   Rs!isRemarksCompulsory = ChkisRemarksCompulsory.Value
   Rs!isKitchen = ChkIsKitchen.Value
   Rs!PCTCode = IIf(Trim(TxtPCTCode.Text) = "", Null, TxtPCTCode.Text)
   Rs.Update
   If vIsNewRecord = True Then Call ActivityLog("Groups", eAdd, , , TxtID.Text)
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
   On Error GoTo ErrorHandler
   If vIsNewRecord Then
      If Trim(TxtID.Text) = "" Then
         MsgBox "Please specify a Group ID", vbExclamation, "Alert"
         If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
         Exit Function
      End If
      If Len(Trim(TxtID.Text)) < 3 Then
         MsgBox "The Group ID must be three characters long", vbExclamation, "Alert"
         TxtID.Text = Right("000" + CStr(Val(TxtID.Text)), 3)
         If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
         Exit Function
      End If
      If cn.Execute("Select count(*) from Groups where Groupid = '" & TxtID.Text & "'").Fields(0) > 0 Then
          MsgBox "This Group ID already exists. The Group ID must be unique", vbExclamation, "Alert"
          TxtID.Text = FunGetMaxID
          If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
          Exit Function
      End If
      Select Case Asc(UCase(Left(TxtID.Text, 1)))
        Case 65 To 90
        Case 48 To 57
        Case Else
          MsgBox "The Group ID must contain numeric/alphabetical characters only", vbExclamation, "Alert"
          If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
          Exit Function
      End Select
      Select Case Asc(UCase(Right(TxtID.Text, 1)))
        Case 65 To 90
        Case 48 To 57
        Case Else
          MsgBox "The Group ID must contain numeric/alphabetical characters only", vbExclamation, "Alert"
          If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
          Exit Function
      End Select
      If cn.Execute("Select * FROM Groups where GroupName like '" & TxtName.Text & "'").RecordCount > 0 Then
         MsgBox "This Group Already Exists", vbExclamation, "Alert"
         If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
         Exit Function
      End If
   Else
      If cn.Execute("Select * FROM Groups where GroupID <> '" & TxtID.Text & "' and GroupName like '" & TxtName.Text & "'").RecordCount > 0 Then
         MsgBox "This Group Already Exists", vbExclamation, "Alert"
         If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
         Exit Function
      End If
   End If
   If Trim(TxtName.Text) = "" Then
      MsgBox "Please specify a Group name", vbExclamation, "Alert"
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
   SetWindowText Me.hWnd, "Groups"
   HelpLocation Me
   LblGroupName1.Visible = ObjRegistry.AllowUrduProduct
   TxtGroupName1.Visible = ObjRegistry.AllowUrduProduct
   Set Rs = New ADODB.Recordset
   Rs.Open "Select * FROM Groups order by GroupName", cn, adOpenDynamic, adLockOptimistic
   Set Grid.DataSource = Rs
   Grid.Columns("ID").DataField = "GroupID"
   Grid.Columns("Name").DataField = "GroupName"
   FormStatus = NewMode
   BtnSave.Visible = Not ObjRegistry.ReadOnlyStatus
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   
   ChkIsKitchen.Visible = ObjRegistry.PrintKitchenInoices
   ModeValue = False
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
         TxtGroupName1.Enabled = False
         TxtID.Enabled = False
         ChkisRemarksCompulsory.Enabled = True
         ChkisRemarksCompulsory.Value = 0
         ChkIsKitchen.Enabled = True
         ChkIsKitchen.Value = 0
         TxtName.Text = ""
         TxtGroupName1.Text = ""
         TxtPCTCode.Text = ""
         TxtID.Text = ""
         TxtFilter.Text = ""
         TxtFilter.Enabled = False
         TxtFilter.BackColor = &HE0E0E0
         TxtID.Text = FunGetMaxID()
         TxtName.Enabled = True
         TxtName.BackColor = vbWhite
         TxtGroupName1.BackColor = vbWhite
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
         TxtGroupName1.Enabled = True
         ChkisRemarksCompulsory.Enabled = True
         ChkIsKitchen.Enabled = True
         TxtID.Enabled = False
         TxtID.BackColor = &HE0E0E0
         TxtPCTCode.Enabled = True
         TxtPCTCode.BackColor = vbWhite
         TxtName.BackColor = vbWhite
         TxtGroupName1.BackColor = vbWhite
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
         TxtGroupName1.Enabled = False
         TxtGroupName1.BackColor = &HE0E0E0
         TxtPCTCode.Enabled = False
         TxtPCTCode.BackColor = &HE0E0E0
         TxtID.Enabled = False
         TxtID.BackColor = &HE0E0E0
         ChkisRemarksCompulsory.Enabled = False
         ChkIsKitchen.Enabled = False
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
      Set DefGroups = Nothing
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
      vStrSQL = "Select isnull(GroupName1,'') from groups where GroupID = '" & TxtID.Text & "'"
      TxtGroupName1.Text = cn.Execute(vStrSQL).Fields(0).Value
      TxtName.Text = Grid.Columns("Name").Text
      TxtPCTCode.Text = IIf(IsNull(Rs!PCTCode), "", Rs!PCTCode)
      ChkisRemarksCompulsory.Value = Abs(Rs!isRemarksCompulsory)
      ChkIsKitchen.Value = Abs(Rs!isKitchen)
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
   If Me.ActiveControl.Name <> TxtFilter.Name Then Exit Sub
   If Trim(TxtFilter.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "GroupName like '" & Replace(TxtFilter.Text, "'", "''") & "%'", , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As String
  On Error GoTo ErrorHandler
  FunGetMaxID = cn.Execute("Select right('000' + cast(isnull(max(GroupID),0) + 1 as varchar) ,3) from Groups").Fields(0)
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub TxtGroupName1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
On Error GoTo ErrorHandler
        ''''''''''''''''''''''''''''''''''''''''''''''''
        ''                                             ''
        '' There are the KeyDown-Event behaviours for  ''
        '' Enter, Space, Delete & Tab keys to set      ''
        '' Behavior in TxtGroupName1.Text, keys will behave ''
        '' as Normal Text writing behavior.            ''
        ''                                             ''
        '''''''''''''''''''''''''''''''''''''''''''''''''

   If ModeValue = False Then
      'Space Key Behavior
         If KeyCode = 32 Then
         UniCode = &H20
         TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)
         KeyCode = 0

        'Enter Key Behavior
'        ElseIf KeyCode = 13 Then
'        UniCode = &HA
'        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)
'        KeyCode = 0

        'Horizontal Tab Behavior
'        ElseIf KeyCode = 9 Then
'        UniCode = &H9
'        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)
'        KeyCode = 0

         'Delete Key Behavior
         ElseIf KeyCode = 127 Then
         UniCode = &H7F
         TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)
         KeyCode = 0
         End If
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
        
        'This Function Got End There
End Sub

Private Sub TxtGroupName1_KeyPress(KeyAscii As MSForms.ReturnInteger)
On Error GoTo ErrorHandler
        ''''''''''''''''''''''''''''''''''''''''''''''''
        ''                                             ''
        '' There are the KeyPress-Event behaviours for ''
        '' Alfabatic, Numaric & Symbolic keys to write ''
        '' Urdu. I've tried to make it near with Urdu  ''
        '' Phonetic Keyboard Layout.                   ''
        ''                                             ''
        '''''''''''''''''''''''''''''''''''''''''''''''''

If ModeValue = False Then

        'For Small Letter's Behaviors

        'a Key Behavior
        If KeyAscii = 97 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H627
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'b Key Behavior
        ElseIf KeyAscii = 98 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H628
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'c Key Behavior
        ElseIf KeyAscii = 99 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H686
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'd Key Behavior
        ElseIf KeyAscii = 100 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H62F
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'e Key Behavior
        ElseIf KeyAscii = 101 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H639
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'f Key Behavior
        ElseIf KeyAscii = 102 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H641
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'g Key Behavior
        ElseIf KeyAscii = 103 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H6AF
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'h Key Behavior
        ElseIf KeyAscii = 104 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H6BE
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'i Key Behavior
        ElseIf KeyAscii = 105 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H6CC
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'j Key Behavior
        ElseIf KeyAscii = 106 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H62C
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'k Key Behavior
        ElseIf KeyAscii = 107 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H6A9
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'l Key Behavior
        ElseIf KeyAscii = 108 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H644
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'm Key Behavior
        ElseIf KeyAscii = 109 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H645
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'n Key Behavior
        ElseIf KeyAscii = 110 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H646
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'o Key Behavior
        ElseIf KeyAscii = 111 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H6C1
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'p Key Behavior
        ElseIf KeyAscii = 112 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H67E
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'q Key Behavior
        ElseIf KeyAscii = 113 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H642
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'r Key Behavior
        ElseIf KeyAscii = 114 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H631
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        's Key Behavior
        ElseIf KeyAscii = 115 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H633
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        't Key Behavior
        ElseIf KeyAscii = 116 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H62A
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'u Key Behavior
        ElseIf KeyAscii = 117 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H621
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'v Key Behavior
        ElseIf KeyAscii = 118 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H637
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'w Key Behavior
        ElseIf KeyAscii = 119 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H648
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'x Key Behavior
        ElseIf KeyAscii = 120 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H634
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'y Key Behavior
        ElseIf KeyAscii = 121 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H6D2
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'z Key Behavior
        ElseIf KeyAscii = 122 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H632
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)


        ' For Capital Latter's Behaviors

        'A Key Behavior
        ElseIf KeyAscii = 65 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H622
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'B Key Behavior
        ElseIf KeyAscii = 66 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &HFBB0
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'C Key Behavior
        ElseIf KeyAscii = 67 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H62B
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'D Key Behavior
        ElseIf KeyAscii = 68 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H688
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'E Key Behavior
        ElseIf KeyAscii = 69 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H650
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'F Key Behavior
        ElseIf KeyAscii = 70 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H652
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'G Key Behavior
        ElseIf KeyAscii = 71 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H63A
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'H Key Behavior
        ElseIf KeyAscii = 72 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H62D
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'I Key Behavior
        ElseIf KeyAscii = 73 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H649
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'J Key Behavior
        ElseIf KeyAscii = 74 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H636
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'K Key Behavior
        ElseIf KeyAscii = 75 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H62E
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'L Key Behavior
        ElseIf KeyAscii = 76 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &HFEFB
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'M Key Behavior
        ElseIf KeyAscii = 77 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H66B
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'N Key Behavior
        ElseIf KeyAscii = 78 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H6BA
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'O Key Behavior
        ElseIf KeyAscii = 79 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H629
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'P Key Behavior
        ElseIf KeyAscii = 80 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H64F
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'Q Key Behavior
        ElseIf KeyAscii = 81 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H626
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'R Key Behavior
        ElseIf KeyAscii = 82 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H691
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'S Key Behavior
        ElseIf KeyAscii = 83 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H635
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'T Key Behavior
        ElseIf KeyAscii = 84 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H679
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'U Key Behavior
        ElseIf KeyAscii = 85 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H626
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'V Key Behavior
        ElseIf KeyAscii = 86 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H638
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'W Key Behavior
        ElseIf KeyAscii = 87 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H624
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'Z Key Behavior
        ElseIf KeyAscii = 88 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H698
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'Y Key Behavior
        ElseIf KeyAscii = 89 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &HFBAF
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        'Z Key Behavior
        ElseIf KeyAscii = 90 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H630
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)


        'For Numaric Key's Behaviors

        '0 Key Behavior
        ElseIf KeyAscii = 48 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H660
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '1 Key Behavior
        ElseIf KeyAscii = 49 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H661
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '2 Key Behavior
        ElseIf KeyAscii = 50 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H662
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '3 Key Behavior
        ElseIf KeyAscii = 51 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H663
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '4 Key Behavior
        ElseIf KeyAscii = 52 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H664
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '5 Key Behavior
        ElseIf KeyAscii = 53 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H665
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '6 Key Behavior
        ElseIf KeyAscii = 54 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H666
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '7 Key Behavior
        ElseIf KeyAscii = 55 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H667
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '8 Key Behavior
        ElseIf KeyAscii = 56 Or TxtGroupName1.SelText <> "" Then
        UniCode = &H668
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '9 Key Behavior
        ElseIf KeyAscii = 57 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H669
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        ' Numaric Keys with 'Shift' Behavior

        ') Key Behavior
        ElseIf KeyAscii = 41 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &HFD3F
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '! Key Behavior
        ElseIf KeyAscii = 33 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H21
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '@ Key Behavior
        ElseIf KeyAscii = 64 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H40
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '# Key Behavior
        ElseIf KeyAscii = 35 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H23
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '$ Key Behavior
        ElseIf KeyAscii = 36 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H24
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '% Key Behavior
        ElseIf KeyAscii = 37 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H66A
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '^ Key Behavior
        ElseIf KeyAscii = 94 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H5E
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '& Key Behavior
        ElseIf KeyAscii = 38 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H26
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '* Key Behavior
        ElseIf KeyAscii = 42 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H66D
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '( Key Behavior
        ElseIf KeyAscii = 40 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &HFD3E
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)


        'For Special Characters

        'Symbols

        '? Key Behavior
        ElseIf KeyAscii = 63 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H61F
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '/ Key Behavior
        ElseIf KeyAscii = 47 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H2F
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        ', Key Behavior
        ElseIf KeyAscii = 44 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H60C
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '. Key Behavior
        ElseIf KeyAscii = 46 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H640
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '_ Key Behavior
        ElseIf KeyAscii = 95 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H5F
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '- Key Behavior
        ElseIf KeyAscii = 45 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H2D
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '+ Key Behavior
        ElseIf KeyAscii = 43 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H2B
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '= Key Behavior
        ElseIf KeyAscii = 61 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H3D
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        ': Key Behavior
        ElseIf KeyAscii = 58 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H3A
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '; Key Behavior
        ElseIf KeyAscii = 59 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H201C
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '< Key Behavior
        ElseIf KeyAscii = 60 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H64E
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '> Key Behavior
        ElseIf KeyAscii = 62 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H650
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '{ Key Behavior
        ElseIf KeyAscii = 123 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H2018
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '} Key Behavior
        ElseIf KeyAscii = 125 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H2019
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '[ Key Behavior
        ElseIf KeyAscii = 91 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H5B
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '] Key Behavior
        ElseIf KeyAscii = 93 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H5D
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '| Key Behavior
        ElseIf KeyAscii = 124 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H7C
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '\ Key Behavior
        ElseIf KeyAscii = 92 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H5C
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '~ Key Behavior
        ElseIf KeyAscii = 126 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H64B
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '` Key Behavior
        ElseIf KeyAscii = 96 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H64D
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '" Key Behavior
        ElseIf KeyAscii = 34 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H2190
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        '' Key Behavior
        ElseIf KeyAscii = 39 Or TxtGroupName1.SelText <> "" Then
        TxtGroupName1.SelText = ""
        UniCode = &H201D
        TxtGroupName1.Text = TxtGroupName1.Text + ChrW(UniCode)

        End If
        KeyAscii = 0
  End If

        'This Function Got End There
Exit Sub
ErrorHandler:
   Call ShowErrorMessage


End Sub




Private Sub TxtName_Change()
   If Me.ActiveControl.Name <> TxtName.Name Then Exit Sub
   TxtFilter.Text = TxtName.Text
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxBinID = cn.Execute("Select isnull(max(BinID),0)+1 from Bin_Groups ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
    If vIsNewRecord = False Then
    With cn.Execute("Select  * from Groups where GroupID =" & TxtID.Text)
        If TxtName.Text <> !GroupName Then
            cn.Execute ("Insert Into UserActivities values ('Groups'" & "," & TxtID.Text & ", Null , 'Updated GroupName-" & !GroupName & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        
    End With
   Else
        cn.Execute ("Insert Into UserActivities values ('Groups'" & "," & TxtID.Text & ", Null ,'Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

