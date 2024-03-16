VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form FrmDesktopSetting 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   5850
      Left            =   3413
      TabIndex        =   0
      Top             =   3113
      Width           =   8595
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   4
      stylesets.count =   2
      stylesets(0).Name=   "SelectedCol"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   12713983
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
      stylesets(0).Picture=   "FrmDesktopSetting.frx":0000
      stylesets(1).Name=   "SelectedRow"
      stylesets(1).ForeColor=   16777215
      stylesets(1).BackColor=   8388608
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "FrmDesktopSetting.frx":001C
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedCol"
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   2
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
      ExtraHeight     =   106
      Columns.Count   =   4
      Columns(0).Width=   7250
      Columns(0).Caption=   "Menu Name"
      Columns(0).Name =   "Name"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   5477
      Columns(1).Caption=   "Caption"
      Columns(1).Name =   "Caption"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1376
      Columns(2).Caption=   "Position"
      Columns(2).Name =   "ID"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Style=   3
      Columns(2).Row.Count=   12
      Columns(2).Col.Count=   2
      Columns(2).Row(0).Col(0)=   "1"
      Columns(2).Row(1).Col(0)=   "2"
      Columns(2).Row(2).Col(0)=   "3"
      Columns(2).Row(3).Col(0)=   "4"
      Columns(2).Row(4).Col(0)=   "5"
      Columns(2).Row(5).Col(0)=   "6"
      Columns(2).Row(6).Col(0)=   "7"
      Columns(2).Row(7).Col(0)=   "8"
      Columns(2).Row(8).Col(0)=   "9"
      Columns(2).Row(9).Col(0)=   "10"
      Columns(2).Row(10).Col(0)=   "11"
      Columns(2).Row(11).Col(0)=   "12"
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "TaskKey"
      Columns(3).Name =   "TaskKey"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   15161
      _ExtentY        =   10319
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
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   6068
      TabIndex        =   1
      Top             =   9398
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
      MICON           =   "FrmDesktopSetting.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7373
      TabIndex        =   2
      Top             =   9398
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
      MICON           =   "FrmDesktopSetting.frx":0054
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8663
      TabIndex        =   3
      Top             =   9398
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
      MICON           =   "FrmDesktopSetting.frx":0070
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desktop Setting"
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
      TabIndex        =   4
      Top             =   270
      Width           =   2190
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11288
      Top             =   2093
      Width           =   330
   End
End
Attribute VB_Name = "FrmDesktopSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim Rs As New ADODB.Recordset
Dim vSuppressUpdateEvent As Boolean
Dim vPosition As Byte
Dim vTaskKey As String
Dim vRowCounter As Integer
Dim vFind As Boolean
Dim vSQL As String
'----------------------------------

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   If Rs.State = adStateOpen Then
      Rs.CancelBatch
      Rs.Close
   End If
   Me.MousePointer = vbHourglass
   Rs.Open "Select * From DesktopShortcuts where userno =" & ObjUserSecurity.UserNo, CN, adOpenStatic, adLockBatchOptimistic
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
   If ObjUserSecurity.IsAdministrator = True Then
      vSQL = "Select t.TaskKey, Description, Caption, isnull(Position,0) as Position FROM Tasks t Left Outer Join (select * from DesktopShortcuts where userno =" & ObjUserSecurity.UserNo & " )s on t.TaskKey = s.TaskKey order by [description]"
   Else
      vSQL = "Select t.TaskKey, Description, Caption, isnull(Position,0) as Position FROM (select * from UserTasks where userno =" & ObjUserSecurity.UserNo & " )u inner join Tasks t on t.TaskKey = u.TaskKey Left Outer Join (select * from DesktopShortcuts where userno =" & ObjUserSecurity.UserNo & " )s on t.TaskKey = s.TaskKey order by [description]"
   End If
   With CN.Execute(vSQL)
      Do Until .EOF
         Grid.AddNew
         Grid.Columns("Taskkey").Text = .Fields("TaskKey").Value
         Grid.Columns("Caption").Text = IIf(IsNull(.Fields("Caption").Value), "", .Fields("Caption").Value)
         Grid.Columns("Name").Text = IIf(IsNull(.Fields("description").Value), "", .Fields("description").Value)
         Grid.Columns("Position").Value = .Fields("Position").Value
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
   LoadGrid
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   Grid.Update
   Rs.Filter = ""
   'Rs.MoveFirst
   'While Not Rs.EOF
   '   If Rs.EditMode <> adEditNone Then
   '      Call ActivityLog("Account Opening Balance", eEdit, , , Rs!AccountNo)
   '   End If
   '  Rs.MoveNext
   'Wend
   Rs.UpdateBatch
   Call Desktop.EnableShortcuts
   MsgBox "Your Desktop Setting has been Changed Successfully.", vbOKOnly + vbInformation, "Information"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "Desktop Setting"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   LoadGrid
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
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnSave.Enabled Then BtnSave_Click
            KeyCode = 0
         Case vbKeyW
            'If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   End If
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
   On Error GoTo ErrorHandler
   If Grid.Columns("Position").Text = "0" Then Grid.Columns("Position").Text = ""
   If vSuppressUpdateEvent Then Exit Sub
   If Val(Grid.Columns("Position").Value) > 12 Then
       'MsgBox "Position Can Not Greater Than in the List"
      Grid.Columns("Position").Value = ""
      Exit Sub
   End If
   vSuppressUpdateEvent = True
   vTaskKey = Grid.Columns("TaskKey").Text
   vPosition = Val(Grid.Columns("Position").Value)
   Dim vCurrentBM As Variant
   Dim vBM As Variant
   vCurrentBM = Grid.Bookmark
   Dim i As Integer
   vFind = False
   Grid.MoveFirst
   For i = 0 To Grid.Rows - 1
      If Grid.Columns(2).CellValue(Grid.GetBookmark(i)) = vPosition And Grid.Columns(3).CellValue(Grid.GetBookmark(i)) <> vTaskKey Then
         vBM = Grid.GetBookmark(i)
         vFind = True
      End If
   Next i
   If vFind = True Then
      Grid.Bookmark = vBM
      Grid.Columns("Position").Value = ""
      Rs.Filter = "TaskKey='" & Grid.Columns("TaskKey").Value & "'"
      Rs.Delete
   End If
   Grid.Bookmark = vCurrentBM
'   If vPosition > 0 Then
'      Grid.Redraw = False
'      vRowCounter = 1
'      Grid.MoveFirst
'      While (vRowCounter <> Grid.Rows) And (Not (Val(Grid.Columns("Position").Value) = vPosition And Grid.Columns("TaskKey").Text <> vTaskKey))
'         vRowCounter = vRowCounter + 1
'         Grid.MoveNext
'      Wend
'      If Val(Grid.Columns("Position").Value) = vPosition And Grid.Columns("TaskKey").Text <> vTaskKey Then
'         Grid.Columns("Position").Value = ""
'         Rs.Filter = "TaskKey='" & Grid.Columns("TaskKey").Value & "'"
'         Rs.Delete
'      End If
'   End If
'   Grid.MoveFirst
'   While Grid.Columns("TaskKey").Text <> vTaskKey
'      Grid.MoveNext
'   Wend
'   'Grid.MoveNext
   vSuppressUpdateEvent = False
   Rs.Filter = "TaskKey='" & Grid.Columns("TaskKey").Value & "'"
   If Rs.RecordCount = 0 And Val(Grid.Columns("Position").Value) > 0 Then
      Rs.AddNew
      Rs!UserNo = ObjUserSecurity.UserNo
      Rs!TaskKey = Grid.Columns("TaskKey").Text
      Rs!Caption = IIf(Grid.Columns("Caption").Text = "", Grid.Columns("Name").Text, Grid.Columns("Caption").Text)
      Rs!Position = Grid.Columns("Position").Value
      vFind = True
   ElseIf Rs.RecordCount = 1 And Val(Grid.Columns("Position").Value) = 0 Then
      'Rs!UserNo = 0
      Rs.Delete
      vFind = True
   ElseIf Rs.RecordCount = 1 And Val(Grid.Columns("Position").Value) > 0 Then
      Rs!TaskKey = Grid.Columns("TaskKey").Text
      Rs!Caption = IIf(Grid.Columns("Caption").Text = "", Grid.Columns("Name").Text, Grid.Columns("Caption").Text)
      Rs!Position = Grid.Columns("Position").Value
      vFind = True
   End If
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Grid.Redraw = True
   Call ShowErrorMessage
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

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   If vFind = True Then Grid.Col = 2: vFind = False
End Sub

Private Sub Grid_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
   Grid.Col = 2
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub
