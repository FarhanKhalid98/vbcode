VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form DefReferences 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "DefReferences.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCNIC 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10665
      MaxLength       =   15
      TabIndex        =   8
      Top             =   7388
      Width           =   2655
   End
   Begin VB.TextBox TxtCast 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   10665
      MaxLength       =   30
      TabIndex        =   6
      Top             =   6743
      Width           =   2595
   End
   Begin VB.TextBox TxtAddress2 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   8010
      MaxLength       =   100
      TabIndex        =   4
      Top             =   5648
      Width           =   5265
   End
   Begin VB.TextBox TxtFName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8010
      MaxLength       =   30
      TabIndex        =   2
      Top             =   3998
      Width           =   5085
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
      Left            =   11250
      TabIndex        =   25
      Top             =   765
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
         TabIndex        =   26
         Tag             =   "NC"
         Text            =   "DefReferences.frx":0ECA
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
         TabIndex        =   27
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.TextBox TxtPrefix 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   8010
      MaxLength       =   2
      TabIndex        =   18
      Tag             =   "NC"
      Top             =   2678
      Width           =   525
   End
   Begin VB.TextBox TxtPhone 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8010
      MaxLength       =   30
      TabIndex        =   7
      Top             =   7388
      Width           =   2655
   End
   Begin VB.TextBox TxtCity 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8010
      MaxLength       =   30
      TabIndex        =   5
      Top             =   6743
      Width           =   2655
   End
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   8010
      MaxLength       =   100
      TabIndex        =   3
      Top             =   4628
      Width           =   5265
   End
   Begin VB.TextBox TxtID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8535
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2678
      Width           =   825
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8010
      MaxLength       =   30
      TabIndex        =   1
      Top             =   3323
      Width           =   5085
   End
   Begin VB.TextBox TxtFilter 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2925
      MaxLength       =   30
      TabIndex        =   12
      Tag             =   "NC"
      Top             =   1943
      Width           =   4395
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5700
      Left            =   2040
      TabIndex        =   13
      Top             =   2273
      Width           =   5295
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
      stylesets(0).Picture=   "DefReferences.frx":0F55
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
      Columns(0).Width=   2037
      Columns(0).Caption=   "Reference ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   6271
      Columns(1).Caption=   "Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   9340
      _ExtentY        =   10054
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
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   2985
      TabIndex        =   14
      Top             =   8783
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
      MICON           =   "DefReferences.frx":0F71
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   4305
      TabIndex        =   15
      Top             =   8783
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
      MICON           =   "DefReferences.frx":0F8D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   5625
      TabIndex        =   16
      Top             =   8783
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
      MICON           =   "DefReferences.frx":0FA9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   8430
      TabIndex        =   9
      Top             =   8783
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
      MICON           =   "DefReferences.frx":0FC5
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   9750
      TabIndex        =   10
      Top             =   8783
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
      MICON           =   "DefReferences.frx":0FE1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   11070
      TabIndex        =   11
      Top             =   8783
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
      MICON           =   "DefReferences.frx":0FFD
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNIC No"
      Height          =   195
      Left            =   10665
      TabIndex        =   32
      Top             =   7178
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cast"
      Height          =   195
      Left            =   10710
      TabIndex        =   31
      Top             =   6518
      Width           =   315
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address 2"
      Height          =   195
      Left            =   8010
      TabIndex        =   30
      Top             =   5438
      Width           =   705
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      Height          =   195
      Left            =   8010
      TabIndex        =   29
      Top             =   3773
      Width           =   915
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
      Left            =   11250
      TabIndex        =   28
      Top             =   495
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "References"
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
      TabIndex        =   24
      Top             =   270
      Width           =   1515
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No"
      Height          =   195
      Left            =   8010
      TabIndex        =   23
      Top             =   7178
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   195
      Left            =   8010
      TabIndex        =   22
      Top             =   6518
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address 1"
      Height          =   195
      Left            =   8010
      TabIndex        =   21
      Top             =   4418
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference ID"
      Height          =   195
      Left            =   8010
      TabIndex        =   20
      Top             =   2468
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   8010
      TabIndex        =   19
      Top             =   3098
      Width           =   420
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   509
      X2              =   508
      Y1              =   128.533
      Y2              =   534.533
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   2310
      TabIndex        =   17
      Top             =   2033
      Width           =   510
   End
End
Attribute VB_Name = "DefReferences"
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   End If
End Sub

Private Sub Grid_DblClick()
   If Grid.Rows > 0 And BtnOpen.Enabled Then BtnOpen_Click
End Sub

Private Sub Grid_GotFocus()
   On Error GoTo ErrorHandler
   Call Grid_RowColChange(0, 0)
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

Private Sub BtnClear_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    CN.Execute ("Insert Into UserActivities values ('Refers'" & "," & TxtID.Text & ",Null,'Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  FormStatus = SelectionMode
End Sub

Private Sub BtnClose_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    CN.Execute ("Insert Into UserActivities values ('Refers'" & "," & TxtID.Text & ",Null,'Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniReferences", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
  Dim vtbl As String
  If Rs.RecordCount > 0 Then
    If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
    Dim vid As String
    vid = Rs!ReferID
    vtbl = Common.ChildDataExists("Refers", "ReferID='" & vid & "'", "")
    If vtbl <> "" Then
      MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
      Exit Sub
    End If
    cn.BeginTrans
    
'    vMaxBinID = FunGetMaxBinID
    ''''''''''''''''''''''''''''''''''''''''''''''''Bin Header----------------------------------------------
'    CN.Execute ("Insert Into Bin_Refers Select " & vMaxBinID & ",'" & Date & "',* from Refers Where ReferID = " & TxtPrefix.Text & TxtID.Text)
    
    Call ActivityLog("Refers", eDelete, , , TxtPrefix.Text & TxtID.Text)
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('Refers'" & "," & TxtPrefix.Text & TxtID.Text & ",Null,'Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Rs.Delete
    'CN.Execute ("Delete From ChartOfAccounts Where AccountNo = '" & vid & "'")
    cn.CommitTrans
    If Rs.RecordCount = 0 Then FormStatus = NewMode: Exit Sub
    Rs.MoveNext
    Grid.MoveNext
    If Rs.EOF Then Rs.MoveLast
  End If
  Exit Sub
ErrorHandler:
  If cn.Errors.Count > 0 Then cn.RollbackTrans
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
   vUserAction = UserAuthentication("MniReferences", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
  If vIsNewRecord = False Then Call ActivityLog("Refers", eEdit, , , TxtID.Text)
  Set Rs = New ADODB.Recordset
  Rs.Open " Select * FROM Refers where ReferID = '" & TxtID.Text & "'", cn, adOpenDynamic, adLockOptimistic
  
  Call UserActivities
  
  cn.BeginTrans
  If vIsNewRecord Then
'   CN.Execute ("Insert into chartofaccounts values (" & _
    "'" & TxtPrefix.Text & TxtID.Text & "',1,'" & Replace(TxtName.Text, "'", "''") & "','Refers',2,'" & Replace(TxtAddress.Text, "'", "''") & "','62',0,0,1," & ChkLockCustomer.Value & ",1,0,' ',0)")
    Rs.AddNew
    Rs!ReferID = TxtID.Text
  Else
'    CN.Execute ("Update Chartofaccounts set Accountname = '" & Replace(TxtName.Text, "'", "''") & "',Narration = '" & Replace(TxtAddress.Text, "'", "''") & "', isLocked = " & ChkLockCustomer.Value & " Where AccountNo = '" & Rs!ReferID & "'")
  End If
  Rs!Name = TxtName.Text
  Rs!FName = TxtFName.Text
  Rs!Address = TxtAddress.Text
  Rs!Address2 = TxtAddress2.Text
  Rs!City = TxtCity.Text
  Rs!Cast = TxtCast.Text
  Rs!Phone = TxtPhone.Text
  Rs!CNIC = TxtCNIC.Text
  Rs.Update
  cn.CommitTrans
  Set Rs = New ADODB.Recordset
  If Rs.State = adStateOpen Then Rs.Close
  Rs.Open "Select * FROM Refers", cn, adOpenDynamic, adLockOptimistic
  If vIsNewRecord = True Then Call ActivityLog("Refers", eAdd, , , TxtPrefix.Text & TxtID.Text)
  FormStatus = NewMode
  Exit Sub
ErrorHandler:
  If cn.Errors.Count > 0 Then cn.RollbackTrans
  Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
  If vIsNewRecord Then
    If Trim(TxtID.Text) = "" Then
      MsgBox "Please specify a Reference ID", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
    If Not IsNumeric(TxtID.Text) Then
      MsgBox "The Reference ID must be numeric", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
  End If
  If Trim(TxtName.Text) = "" Then
    MsgBox "Please specify a Reference Name", vbExclamation, "Alert"
    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
    Exit Function
  End If
  If TxtID.Enabled = True And cn.Execute("select count(*) from Refers where ReferID = '" & TxtID.Text & "'").Fields(0) > 0 Then
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
   SetWindowText Me.hWnd, "Refers"
   HelpLocation Me
   Set Rs = New ADODB.Recordset
   Rs.Open "Select * FROM Refers", cn, adOpenDynamic, adLockOptimistic
   Grid.Columns("ID").DataField = "ReferID"
   Grid.Columns("Name").DataField = "Name"
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
      Call SubClearFields(True)
      BtnNew.Enabled = False
      BtnOpen.Enabled = False
      BtnDelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtPrefix.Enabled = False
      TxtPrefix.Text = "64"
      TxtID.Text = FunGetMaxID
      TxtFilter.Text = ""
      Grid.Enabled = False
      Set Grid.DataSource = Rs
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
      TxtFilter.Text = ""
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
      'TxtFilter.Text = ""
  End Select
  Exit Property
ErrorHandler:
  Call ShowErrorMessage
End Property

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo ErrorHandler
   If Rs.RecordCount > 0 And Grid.Enabled Then
      'Rs!Name = TxtName.Text
      TxtID.Text = Val(Grid.Columns("ID").Text)
      TxtName.Text = Rs!Name
      TxtFName.Text = IIf(IsNull(Rs!FName), "", Rs!FName)
      TxtAddress.Text = IIf(IsNull(Rs!Address), "", Rs!Address)
      TxtAddress2.Text = IIf(IsNull(Rs!Address2), "", Rs!Address2)
      TxtCity.Text = IIf(IsNull(Rs!City), "", Rs!City)
      TxtCast.Text = IIf(IsNull(Rs!Cast), "", Rs!Cast)
      TxtPhone.Text = IIf(IsNull(Rs!Phone), "", Rs!Phone)
      TxtCNIC.Text = IIf(IsNull(Rs!CNIC), "", Rs!CNIC)
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
   FunGetMaxID = cn.Execute("Select isnull(max(ReferID),0) + 1 from Refers").Fields(0)
End Function

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub TxtFilter_Change()
   On Error GoTo ErrorHandler
   'If Me.ActiveControl.Name <> TxtFilter.Name Then Exit Sub
   'If Trim(TxtFilter.Text) = "" Then Grid.MoveFirst: Exit Sub
   Set Rs = New ADODB.Recordset
   Rs.Open "Select * FROM Refers where Name like '%" & Replace(TxtFilter.Text, "'", "''") & "%' Order by Name", cn, adOpenDynamic, adLockOptimistic
   Set Grid.DataSource = Rs
   'Rs.Find "ReferenceName like '" & Replace(TxtFilter.Text, "'", "''") & "%'", , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtPhone_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtCNIC_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9, Asc("/"), Asc("-"), Asc(" "), vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub TxtName_Change()
   If Me.ActiveControl.Name <> TxtName.Name Then Exit Sub
   TxtFilter.Text = TxtName.Text
End Sub

Private Sub TxtName_LostFocus()
   TxtName.Text = StrConv(TxtName.Text, vbProperCase)
End Sub

Private Sub TxtAddress_LostFocus()
   TxtAddress.Text = StrConv(TxtAddress.Text, vbProperCase)
End Sub

Private Sub TxtAddress2_LostFocus()
   TxtAddress2.Text = StrConv(TxtAddress2.Text, vbProperCase)
End Sub

Private Sub TxtFName_LostFocus()
   TxtFName.Text = StrConv(TxtFName.Text, vbProperCase)
End Sub

Private Sub TxtCast_LostFocus()
   TxtCast.Text = StrConv(TxtCast.Text, vbProperCase)
End Sub

Private Sub TxtCity_LostFocus()
   TxtCity.Text = StrConv(TxtCity.Text, vbProperCase)
End Sub

Private Function FunGetMaxBinID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxBinID = cn.Execute("Select isnull(max(BinID),0)+1 from Bin_Refers").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub UserActivities()
    If vIsNewRecord = False Then
        If TxtName.Text <> IIf(IsNull(Rs!Name), "", Rs!Name) Then
            cn.Execute ("Insert Into UserActivities values ('Refers'" & "," & TxtID.Text & ", Null , 'Updated Customer-" & Rs!ReferenceName & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtAddress.Text <> IIf(IsNull(Rs!Address), "", Rs!Address) Then
            cn.Execute ("Insert Into UserActivities values ('Refers'" & "," & TxtID.Text & ", Null , 'Updated Address-" & Rs!Address & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtCity.Text <> IIf(IsNull(Rs!City), "", Rs!City) Then
            cn.Execute ("Insert Into UserActivities values ('Refers'" & "," & TxtID.Text & ", Null , 'Updated City-" & Rs!City & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtPhone.Text <> IIf(IsNull(Rs!Phone), "", Rs!Phone) Then
            cn.Execute ("Insert Into UserActivities values ('Refers'" & "," & TxtID.Text & ", Null , 'Updated Phone1-" & Rs!Phone1 & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
        If TxtCNIC.Text <> IIf(IsNull(Rs!CNIC), "", Rs!CNIC) Then
            cn.Execute ("Insert Into UserActivities values ('Refers'" & "," & TxtID.Text & ", Null , 'Updated Mobile-" & Rs!Mobile & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
   Else
        cn.Execute ("Insert Into UserActivities values ('Refers'" & "," & TxtID.Text & ", Null ,'Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub

