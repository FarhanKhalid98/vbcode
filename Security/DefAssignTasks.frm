VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form DefAssignTasks 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkDeleteAll 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "All"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11175
      TabIndex        =   9
      Top             =   1890
      Width           =   645
   End
   Begin VB.CheckBox ChkEditAll 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "All"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10365
      TabIndex        =   8
      Top             =   1890
      Width           =   645
   End
   Begin VB.CheckBox ChkAllowedAll 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "All"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9375
      TabIndex        =   7
      Top             =   1890
      Width           =   645
   End
   Begin VB.ComboBox CmbFilter 
      Height          =   315
      Left            =   5003
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1703
      Width           =   3150
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6600
      Left            =   2970
      TabIndex        =   1
      Top             =   2130
      Width           =   9420
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   5
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
      stylesets(0).Picture=   "DefAssignTasks.frx":0000
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
      Columns.Count   =   5
      Columns(0).Width=   10557
      Columns(0).Caption=   "Task"
      Columns(0).Name =   "Task"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "TaskKey"
      Columns(1).Name =   "TaskKey"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1799
      Columns(2).Caption=   "Allowed"
      Columns(2).Name =   "Allowed"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Style=   2
      Columns(3).Width=   1482
      Columns(3).Caption=   "Edit"
      Columns(3).Name =   "Edit"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Style=   2
      Columns(4).Width=   1588
      Columns(4).Caption=   "Delete"
      Columns(4).Name =   "Delete"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Style=   2
      _ExtentX        =   16616
      _ExtentY        =   11642
      _StockProps     =   79
      Caption         =   "Assigned Tasks"
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
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   5719
      TabIndex        =   3
      Top             =   9008
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
      MICON           =   "DefAssignTasks.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7046
      TabIndex        =   4
      Top             =   9008
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Reset"
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
      MICON           =   "DefAssignTasks.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8366
      TabIndex        =   5
      Top             =   9008
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
      MICON           =   "DefAssignTasks.frx":0054
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Assign Task"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   6
      Top             =   270
      Width           =   2145
   End
   Begin VB.Image ImgExit 
      Height          =   360
      Left            =   11625
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   225
      Left            =   4095
      TabIndex        =   2
      Top             =   1793
      Width           =   900
   End
End
Attribute VB_Name = "DefAssignTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vSuppressUpdateEvent As Boolean
Dim i As Integer
Public ParaInUserNo As Integer
Dim vSql As String


Private Sub ChkAllowedAll_Click()
   Grid.Redraw = False
   Grid.MoveFirst
      For i = 0 To Grid.Rows
         Grid.Columns("Allowed").Value = ChkAllowedAll.Value
         Call DataManipulate
         Grid.MoveNext
      Next
   Grid.MoveFirst
   Grid.Redraw = True
End Sub

Private Sub ChkDeleteAll_Click()
Grid.Redraw = False
   Grid.MoveFirst
      For i = 0 To Grid.Rows
         Grid.Columns("Delete").Value = ChkDeleteAll.Value
         Call DataManipulate
         Grid.MoveNext
      Next
   Grid.MoveFirst
   Grid.Redraw = True
End Sub

Private Sub ChkEditAll_Click()
Grid.Redraw = False
   Grid.MoveFirst
      For i = 0 To Grid.Rows
         Grid.Columns("Edit").Value = ChkEditAll.Value
         Call DataManipulate
         Grid.MoveNext
      Next
   Grid.MoveFirst
   Grid.Redraw = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name <> Grid.Name Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
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
      End Select
   End If
End Sub

Private Sub CmbFilter_Click()
   On Error GoTo ErrorHandler
   If Rs.State = adStateOpen Then
      Rs.CancelBatch
      Rs.Close
   End If
   If CmbFilter.ListIndex = -1 Then Exit Sub
   Me.MousePointer = vbHourglass
   Rs.Open "Select * From UserTasks where userno = " & Val(CmbFilter.ItemData(CmbFilter.ListIndex)), CN, adOpenStatic, adLockBatchOptimistic
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
'   vSql = "Select distinct taskkey,description,max(allowed) as allowed, isUserEdit, isUserDelete from (Select Tasks.TaskKey,Description,1 as Allowed, isUserEdit, isUserDelete  FROM UserTasks Inner Join Tasks on Tasks.TaskKey = UserTasks.TaskKey Where islocked = 0 and Userno = " & Val(CmbFilter.ItemData(CmbFilter.ListIndex)) & _
'         " UNION " & _
'         " Select TaskKey,Description,0,1,1 from tasks where islocked = 0 and TaskKey like 'mni%') as Data Group by TaskKey,Description, isUserEdit, isUserDelete"
'   With CN.Execute(vSql)
'      Do Until .EOF
'         Grid.AddNew
'         Grid.Columns("Taskkey").Text = .Fields("TaskKey").Value
'         Grid.Columns("Task").Text = .Fields("description").Value
'         Grid.Columns("Allowed").Value = .Fields("Allowed").Value
'         Grid.Columns("Edit").Value = .Fields("isUserDelete").Value
'         Grid.Columns("Delete").Value = .Fields("isUserDelete").Value
'         Grid.Update
'         .MoveNext
'      Loop
'   End With
   
'  vSql = "Select ut.*, description, islocked from userTasks ut inner join Tasks t on t.TaskKey = ut.TaskKey Where  Userno = " & Val(CmbFilter.ItemData(CmbFilter.ListIndex))
   vSql = "Select t.*, Allowed, isUserEdit, isUserDelete from Tasks t " & vbCrLf _
         + "left outer join (Select *, 1 as Allowed  from UserTasks where Userno = " & Val(CmbFilter.ItemData(CmbFilter.ListIndex)) & " ) ut " & vbCrLf _
         + "on t.TaskKey = ut.TaskKey where t.TaskKey like 'mni%' and islocked = 0 order by description "


         
   With CN.Execute(vSql)
      Do Until .EOF
         Grid.AddNew
         Grid.Columns("Taskkey").Text = .Fields("TaskKey").Value
         Grid.Columns("Task").Text = .Fields("description").Value
         Grid.Columns("Allowed").Value = IIf(IsNull(.Fields("Allowed").Value), 0, .Fields("Allowed").Value)
          Grid.Columns("Edit").Value = IIf(IsNull(.Fields("isUserEdit").Value), 0, .Fields("isUserEdit").Value)
         Grid.Columns("Delete").Value = IIf(IsNull(.Fields("isUserDelete").Value), 0, .Fields("isUserDelete").Value)
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
   CmbFilter_Click
End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   Dim vSql As String
   Grid.Update
   Rs.Filter = ""
   Rs.MoveFirst
   While Not Rs.EOF
      If Rs.EditMode = adEditAdd Then
         vSql = "INSERT into ActivityLog(userno,FormType,EntryDate,Description,isnew,isedit,isdelete) values(" & Me.ParaInUserNo & ",'Assign Tasks',getdate(),'User No = " & Rs!UserNo & ", User Name = " & CmbFilter.Text & ", TaskKey = " & Rs!TaskKey & "',1,0,0)"
         CN.Execute vSql
      ElseIf Rs.EditMode = adEditInProgress Then
         If Rs!UserNo = 0 Then
            CN.Execute "Exec ProdActivityLog 'Assign Tasks'," & Me.ParaInUserNo & ",3," & CmbFilter.ItemData(CmbFilter.ListIndex) & ",'01-01-1900','" & Rs!TaskKey & "'"
            Rs.Delete
         End If
      End If
      Rs.MoveNext
   Wend
   Rs.UpdateBatch
   MsgBox "Setting has been Changed Successfully.", vbOKOnly + vbInformation, "Information"
   Exit Sub
ErrorHandler:
  Call ShowErrorMessage
'/**************************************/
'  On Error GoTo ErrorHandler
'  Dim vCounter As Integer
'  Grid.Redraw = False
'  Grid.MoveFirst
'  CN.BeginTrans
'  CN.Execute ("Delete From UserTasks where userno = " & Val(CmbFilter.ItemData(CmbFilter.ListIndex)))
'  If Rs.State = adStateOpen Then Rs.Close
'  Rs.Open "Select * FROM UserTasks Where Userno is null", CN, adOpenDynamic, adLockBatchOptimistic
'  For vCounter = 1 To Grid.Rows
'    If Grid.Columns("Allowed").Value = True Then
'      Rs.AddNew
'      Rs!UserNo = CmbFilter.ItemData(CmbFilter.ListIndex)
'      Rs!TaskKey = Grid.Columns("TaskKey").Text
'      Rs.Update
'    End If
'    Grid.MoveNext
'  Next
'  Rs.UpdateBatch
'  CN.CommitTrans
'  Grid.Redraw = True
'  FormStatus = NewMode
'  Exit Sub
'ErrorHandler:
'  Grid.Redraw = True
'  If CN.Errors.Count > 0 Then CN.RollbackTrans
'  Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hwnd, "Assign Tasks"
   With CN.Execute("Select * FROM Users where userno<>1")
      Do Until .EOF
          CmbFilter.AddItem !UserName
          CmbFilter.ItemData(CmbFilter.NewIndex) = !UserNo
          .MoveNext
      Loop
    End With
    If CmbFilter.ListCount > 0 Then CmbFilter.ListIndex = 0
    'FormStatus = NewMode
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
  Call DataManipulate
End Sub

Private Sub Grid_Change()
   If BtnSave.Enabled = False Then BtnSave.Enabled = True
End Sub

Private Sub Grid_GotFocus()
   Grid.Row = 0
   Grid.Col = 0
'   SendKeys "{Right}"
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

Private Sub CheckAll()
   Grid.Redraw = False
   Grid.MoveFirst
   If ChkAllowedAll.Value = 1 Then
      For i = 0 To Grid.Rows
         Grid.Columns("Allowed").Value = 1
         Grid.MoveNext
      Next
   End If
   Grid.MoveFirst
   Grid.Redraw = True
End Sub
Private Sub DataManipulate()
 On Error GoTo ErrorHandler
   If vSuppressUpdateEvent Then Exit Sub
   Rs.Filter = "Taskkey = '" & Grid.Columns("Taskkey").Value & "'"
   If Rs.RecordCount = 0 Then
      Rs.AddNew
      Rs!UserNo = CmbFilter.ItemData(CmbFilter.ListIndex)
      Rs!TaskKey = Grid.Columns("Taskkey").Text
      Rs!isUserEdit = Grid.Columns("Edit").Text
      Rs!isUserDelete = Grid.Columns("Delete").Text
   ElseIf Rs.RecordCount = 1 And Grid.Columns("Allowed").Value = True Then
      Rs!isUserEdit = Grid.Columns("Edit").Value
      Rs!isUserDelete = Grid.Columns("Delete").Value
      Rs.Update
   ElseIf Rs.RecordCount = 1 And Grid.Columns("Allowed").Value = False Then
      Rs.Delete
'   ElseIf Rs.RecordCount = 1 And Grid.Columns("Allowed").Value = True Then
'      Rs!UserNo = CmbFilter.ItemData(CmbFilter.ListIndex)
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
