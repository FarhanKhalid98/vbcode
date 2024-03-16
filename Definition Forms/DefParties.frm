VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form DefParties 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12015
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DefParties.frx":0000
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   801
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtContactPerson 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6300
      MaxLength       =   30
      TabIndex        =   9
      Top             =   6675
      Width           =   5265
   End
   Begin VB.TextBox TxtFax 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8970
      MaxLength       =   30
      TabIndex        =   7
      Top             =   5355
      Width           =   2595
   End
   Begin VB.TextBox TxtPhone2 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8970
      MaxLength       =   30
      TabIndex        =   5
      Top             =   4695
      Width           =   2595
   End
   Begin VB.TextBox TxtPrefix 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   6300
      MaxLength       =   2
      TabIndex        =   19
      Tag             =   "NC"
      Top             =   1695
      Width           =   525
   End
   Begin VB.TextBox TxtEmail 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6300
      MaxLength       =   50
      TabIndex        =   8
      Top             =   6015
      Width           =   5265
   End
   Begin VB.TextBox TxtMobileNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6300
      MaxLength       =   25
      TabIndex        =   6
      Top             =   5355
      Width           =   2655
   End
   Begin VB.TextBox TxtPhone1 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6300
      MaxLength       =   30
      TabIndex        =   4
      Top             =   4695
      Width           =   2655
   End
   Begin VB.TextBox TxtCity 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6300
      MaxLength       =   30
      TabIndex        =   3
      Top             =   4050
      Width           =   5265
   End
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   6300
      MaxLength       =   100
      TabIndex        =   2
      Top             =   3060
      Width           =   5265
   End
   Begin VB.TextBox TxtID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6840
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1695
      Width           =   825
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6300
      MaxLength       =   30
      TabIndex        =   1
      Top             =   2340
      Width           =   5265
   End
   Begin VB.TextBox TxtFilter 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1215
      MaxLength       =   30
      TabIndex        =   13
      Tag             =   "NC"
      Top             =   1335
      Width           =   4395
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5700
      Left            =   330
      TabIndex        =   14
      Top             =   1695
      Width           =   5280
      ScrollBars      =   2
      _Version        =   196616
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
      stylesets(0).ForeColor=   -2147483634
      stylesets(0).BackColor=   -2147483635
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
      stylesets(0).Picture=   "DefParties.frx":E7F2
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
      Columns(0).Caption=   "Party ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   6456
      Columns(1).Caption=   "Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   9313
      _ExtentY        =   10054
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
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   1275
      TabIndex        =   15
      Top             =   7800
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "New"
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
      MICON           =   "DefParties.frx":E80E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   2595
      TabIndex        =   16
      Top             =   7800
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Change"
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
      MICON           =   "DefParties.frx":E82A
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   3915
      TabIndex        =   17
      Top             =   7800
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
      MICON           =   "DefParties.frx":E846
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6720
      TabIndex        =   10
      Top             =   7800
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
      MICON           =   "DefParties.frx":E862
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   8040
      TabIndex        =   11
      Top             =   7800
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
      MICON           =   "DefParties.frx":E87E
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   9360
      TabIndex        =   12
      Top             =   7800
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
      MICON           =   "DefParties.frx":E89A
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      Height          =   195
      Left            =   8970
      TabIndex        =   28
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
      Height          =   195
      Left            =   6300
      TabIndex        =   27
      Top             =   6465
      Width           =   1095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      Height          =   195
      Left            =   6300
      TabIndex        =   26
      Top             =   5805
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No"
      Height          =   195
      Left            =   6300
      TabIndex        =   25
      Top             =   5145
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Nos"
      Height          =   195
      Left            =   6300
      TabIndex        =   24
      Top             =   4485
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Left            =   6300
      TabIndex        =   23
      Top             =   3825
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   6300
      TabIndex        =   22
      Top             =   2850
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party ID"
      Height          =   195
      Left            =   6300
      TabIndex        =   21
      Top             =   1485
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   6300
      TabIndex        =   20
      Top             =   2130
      Width           =   420
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   395
      X2              =   394
      Y1              =   88
      Y2              =   494
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   600
      TabIndex        =   18
      Top             =   1425
      Width           =   510
   End
End
Attribute VB_Name = "DefParties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vMode As FormMode
Dim vMaxBinID As Integer
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
     If ActiveControl.Name = Grid.Name Then Call Grid_DblClick: Exit Sub
     keybd_event 9, 1, 1, 1
     KeyCode = 0
   End If
   If Shift = vbCtrlMask Then
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
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   End If
End Sub

Private Sub Grid_DblClick()
   If Grid.Rows > 0 And BtnOpen.Enabled Then BtnOpen_Click
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKeyA To vbKeyZ, vbKey0 To vbKey9, Asc("a") To Asc("z")
      TxtFilter.Text = Chr(KeyAscii): TxtFilter.SelStart = Len(TxtFilter.Text):  TxtFilter.SetFocus
   End Select
End Sub

Private Sub BtnClear_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CN.Execute ("Insert Into UserActivities values ('Parties'" & "," & TxtID.Text & ",Null,'Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  FormStatus = SelectionMode
End Sub

Private Sub BtnClose_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CN.Execute ("Insert Into UserActivities values ('Parties'" & "," & TxtID.Text & ",Null,'Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  Dim vtbl As String
  If Rs.RecordCount > 0 Then
    If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
    Dim vid As String
    vid = Rs!PartyID
    vtbl = Common.ChildDataExists("Parties", "PartyId='" & vid & "'", "") & Common.ChildDataExists("ChartoFAccounts", "AccountNo='" & vid & "'", "Parties")
    If vtbl <> "" Then
      MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
      Exit Sub
    End If
    CN.BeginTrans
    
    vMaxBinID = FunGetMaxBinID
    ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header----------------------------------------------
    CN.Execute ("Insert Into Bin_Parties Select " & vMaxBinID & ",'" & Date & "',* from Parties Where PartyID = " & TxtPrefix.Text & TxtID.Text)
   
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CN.Execute ("Insert Into UserActivities values ('Parties'" & "," & TxtPrefix.Text & TxtID.Text & ",Null,'Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Rs.Delete
    CN.Execute ("Delete From ChartOfAccounts Where AccountNo = '" & vid & "'")
    CN.CommitTrans
    If Rs.RecordCount = 0 Then FormStatus = NewMode: Exit Sub
    Rs.MoveNext
    Grid.MoveNext
    If Rs.EOF Then Rs.MoveLast
  End If
  Exit Sub
ErrorHandler:
  If CN.Errors.Count > 0 Then CN.RollbackTrans
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
  
  
  Call UserActivities
  
  If vIsNewRecord Then
    CN.BeginTrans
    
    CN.Execute ("Insert into chartofaccounts values (" & _
    "'" & TxtPrefix.Text & TxtID.Text & "',1,'" & Replace(TxtName.Text, "'", "''") & "','Party',2,'" & Replace(TxtAddress.Text, "'", "''") & "','61',0,0,1,0,1,0,' ',0)")
    Rs.AddNew
    Rs!PartyID = TxtPrefix.Text & TxtID.Text
  Else
    CN.BeginTrans
    CN.Execute ("Update Chartofaccounts set Accountname = '" & Replace(TxtName.Text, "'", "''") & "',Narration = '" & Replace(TxtAddress.Text, "'", "''") & "' Where AccountNo = '" & Rs!PartyID & "'")
  End If
  Rs!PartyName = TxtName.Text
  Rs!Address = TxtAddress.Text
  Rs!City = TxtCity.Text
  Rs!phone1 = TxtPhone1.Text
  Rs!phone2 = TxtPhone2.Text
  Rs!Mobile = TxtMobileNo.Text
  Rs!Fax = TxtFax.Text
  Rs!Email = TxtEmail.Text
  Rs!contactperson = TxtContactPerson.Text
  Rs!PartyType = "V"
  Rs.Update
  CN.CommitTrans
  Rs.Requery
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  If CN.Errors.Count > 0 Then CN.RollbackTrans
  Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
  If vIsNewRecord Then
    If Trim(TxtID.Text) = "" Then
      MsgBox "Please specify a Party ID", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
    If Not IsNumeric(TxtID.Text) Then
      MsgBox "The Party ID must be numeric", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
  End If
  If Trim(TxtName.Text) = "" Then
    MsgBox "Please specify a Party Name", vbExclamation, "Alert"
    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
    Exit Function
  End If
  If TxtID.Enabled = True And CN.Execute("select count(*) from chartofaccounts where accountno = '" & TxtPrefix.Text & TxtID.Text & "'").Fields(0) > 0 Then
    MsgBox "This ID already exists. A new ID has been generated. Please save again", vbExclamation, "Alert"
    TxtID.Text = FunGetMaxID
    TxtID.SetFocus
    Exit Function
  End If
  'All Ok, now validation is success
  FunValidation = True
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
   If UCase(Me.ActiveControl.Name) Like "TXT*" And ActiveControl.Tag = "" Then FormStatus = ChangeMode
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * FROM Parties Where PartyType='V'", CN, adOpenDynamic, adLockOptimistic
    Set Grid.DataSource = Rs
    Grid.Columns("ID").DataField = "PartyId"
    Grid.Columns("Name").DataField = "Partyname"
    FormStatus = NewMode
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
      Call SubClearFields(True)
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtPrefix.Enabled = False
      TxtPrefix.Text = "61"
      TxtID.Text = FunGetMaxID
      Grid.Enabled = False
      TxtFilter.Enabled = False
      If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
      vIsNewRecord = True
    Case Is = OpenMode
      Call SubClearFields(True)
      Call Grid_RowColChange(0, 0)
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnClear.Enabled = True
      Grid.Enabled = False
      TxtID.Enabled = False
      TxtFilter.Enabled = False
      TxtName.SetFocus
      vIsNewRecord = False
    Case Is = ChangeMode
      BtnSave.Enabled = True
    Case Is = SelectionMode
      Grid.Enabled = True
      Call SubClearFields(False)
      Call Grid_RowColChange(0, 0)
      TxtFilter.Enabled = True
      BtnNew.Enabled = True
      BtnOpen.Enabled = True
      BtnDelete.Enabled = True
      BtnSave.Enabled = False
      BtnClear.Enabled = False
      Grid.SetFocus
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo ErrorHandler
  If Rs.RecordCount > 0 And Grid.Enabled Then
    TxtID.Text = Mid(Grid.Columns("ID").Text, 3)
    TxtName.Text = Grid.Columns("Name").Text
    TxtAddress.Text = Rs!Address
    TxtCity.Text = Rs!City
    TxtPhone1.Text = Rs!phone1
    TxtPhone2.Text = Rs!phone2
    TxtMobileNo.Text = Rs!Mobile
    TxtFax.Text = Rs!Fax
    TxtEmail.Text = Rs!Email
    TxtContactPerson.Text = Rs!contactperson
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub SubClearFields(Enable As Boolean)
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is TextBox Then
         If ctl.Tag = "" Then
            ctl.Text = ""
            ctl.Enabled = Enable
         End If
      ElseIf TypeOf ctl Is ComboBox Then
         ctl.Enabled = Enable
      End If
   Next
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As String
  FunGetMaxID = CN.Execute("Select isnull(max(cast(substring(accountno,3,10) as int)),0) + 1 from chartofaccounts Where AccountNo like '61%' and isdetailed=1").Fields(0)
End Function

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtAddress_LostFocus()
   TxtAddress.Text = StrConv(TxtAddress.Text, vbProperCase)
End Sub

Private Sub TxtCity_LostFocus()
   TxtCity.Text = StrConv(TxtCity.Text, vbProperCase)
End Sub

Private Sub TxtContactPerson_LostFocus()
   TxtContactPerson.Text = StrConv(TxtContactPerson.Text, vbProperCase)
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, vbKeyA To vbKeyZ, Asc("@"), Asc("_"), Asc("-"), Asc(" "), vbKeyBack, Asc("a") To Asc("z"), Asc(".")
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtFax_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtFilter_Change()
  On Error GoTo ErrorHandler
  If Trim(TxtFilter.Text) = "" Then Grid.MoveFirst: Exit Sub
  Rs.Find "PartyName like '" & TxtFilter.Text & "%'", , adSearchForward, 1
  If Rs.EOF Then Grid.MoveLast
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub TxtMobileNo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtName_LostFocus()
   TxtName.Text = StrConv(TxtName.Text, vbProperCase)
End Sub

Private Sub TxtPhone1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtPhone2_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxBinID = CN.Execute("Select isnull(max(BinID),0)+1 from Bin_Parties ").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
    If vIsNewRecord = False Then
        If TxtName.Text <> IIf(IsNull(Rs!PartyName), "", Rs!PartyName) Then
            CN.Execute ("Insert Into UserActivities values ('Parties'" & "," & TxtID.Text & ", Null , 'Updated Party Name-" & Rs!PartyName & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtAddress.Text <> IIf(IsNull(Rs!Address), "", Rs!Address) Then
            CN.Execute ("Insert Into UserActivities values ('Parties'" & "," & TxtID.Text & ", Null , 'Updated Address-" & Rs!Address & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtCity.Text <> IIf(IsNull(Rs!City), "", Rs!City) Then
            CN.Execute ("Insert Into UserActivities values ('Parties'" & "," & TxtID.Text & ", Null , 'Updated City-" & Rs!City & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtPhone1.Text <> IIf(IsNull(Rs!phone1), "", Rs!phone1) Then
            CN.Execute ("Insert Into UserActivities values ('Parties'" & "," & TxtID.Text & ", Null , 'Updated Phone1-" & Rs!phone1 & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtPhone2.Text <> IIf(IsNull(Rs!phone2), "", Rs!phone2) Then
            CN.Execute ("Insert Into UserActivities values ('Parties'" & "," & TxtID.Text & ", Null , 'Updated Phone2-" & Rs!phone2 & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtMobileNo.Text <> IIf(IsNull(Rs!Mobile), "", Rs!Mobile) Then
            CN.Execute ("Insert Into UserActivities values ('Parties'" & "," & TxtID.Text & ", Null , 'Updated Mobile-" & Rs!Mobile & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtFax.Text <> IIf(IsNull(Rs!Fax), "", Rs!Fax) Then
            CN.Execute ("Insert Into UserActivities values ('Parties'" & "," & TxtID.Text & ", Null , 'Updated Fax-" & Rs!Fax & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtEmail.Text <> IIf(IsNull(Rs!Email), "", Rs!Email) Then
            CN.Execute ("Insert Into UserActivities values ('Parties'" & "," & TxtID.Text & ", Null , 'Updated Email-" & Rs!Email & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtContactPerson.Text <> IIf(IsNull(Rs!contactperson), "", Rs!contactperson) Then
            CN.Execute ("Insert Into UserActivities values ('Parties'" & "," & TxtID.Text & ", Null , 'Updated ContactPerson-" & Rs!contactperson & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
   Else
        CN.Execute ("Insert Into UserActivities values ('Parties'" & "," & TxtID.Text & ", Null ,'Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub


