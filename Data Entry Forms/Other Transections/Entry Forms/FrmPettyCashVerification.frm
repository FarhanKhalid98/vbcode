VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmPettyCashVerification 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmPettyCashVerification.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnVerified 
      Height          =   420
      Left            =   7710
      TabIndex        =   0
      Top             =   8040
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Verify"
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
      MICON           =   "FrmPettyCashVerification.frx":0ECA
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtAmount 
      Height          =   315
      Left            =   9510
      TabIndex        =   2
      Top             =   7320
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   7
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
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   3735
      Left            =   5820
      TabIndex        =   1
      Top             =   3540
      Width           =   4890
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   5
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
      stylesets(0).Picture=   "FrmPettyCashVerification.frx":0EE6
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
      stylesets(1).Picture=   "FrmPettyCashVerification.frx":0F02
      AllowUpdate     =   0   'False
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
      ExtraHeight     =   106
      Columns.Count   =   5
      Columns(0).Width=   2117
      Columns(0).Caption=   "Denomination"
      Columns(0).Name =   "Denom"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   714
      Columns(1).Name =   "Mul"
      Columns(1).Alignment=   2
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1296
      Columns(2).Caption=   "Quantity"
      Columns(2).Name =   "Qty"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   794
      Columns(3).Name =   "Equ"
      Columns(3).Alignment=   2
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   2646
      Columns(4).Caption=   "Amount"
      Columns(4).Name =   "Amount"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      TabNavigation   =   1
      _ExtentX        =   8625
      _ExtentY        =   6588
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
   Begin SITextBox.Txt TxtID 
      Height          =   315
      Left            =   4650
      TabIndex        =   5
      Top             =   2955
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   9
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
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4650
      TabIndex        =   6
      Top             =   2670
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8070
      TabIndex        =   4
      Top             =   7365
      Width           =   1380
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Petty Cash Verification"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   2700
      TabIndex        =   3
      Top             =   270
      Width           =   3150
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11610
      Top             =   60
      Width           =   375
   End
End
Attribute VB_Name = "FrmPettyCashVerification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim RsBody As New ADODB.Recordset
Dim vMode As FormMode
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.
Dim vid As String
Dim sSql As String, vStrSQL As String, vCounter As Integer

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   Me.MousePointer = vbHourglass
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   sSql = " Select d.Denom, isnull(Qty,0) Qty  " & vbCrLf _
      + " FROM (select * from PettyCashHeader where ID = " & Val(TxtID.Text) & ")h " & vbCrLf _
      + " inner join PettyCashBody b on h.ID = b.ID " & vbCrLf _
      + " right Outer Join Denominations d on b.Denom = d.Denom " & vbCrLf _
      + " Order by d.Denom desc"
   TxtAmount.Text = "0"
   With CN.Execute(sSql)
      Do Until .EOF
         Grid.AddNew
         Grid.Columns("Denom").Value = .Fields("Denom").Value
         Grid.Columns("Mul").Text = "X"
         Grid.Columns("Equ").Text = "="
         Grid.Columns("Qty").Value = .Fields("Qty").Value
         Grid.Columns("Amount").Value = Val(.Fields("Denom").Value) * (.Fields("Qty").Value)
         TxtAmount.Text = Val(TxtAmount.Text) + Grid.Columns("Amount").Value
         Grid.Update
         .MoveNext
      Loop
   End With
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
   On Error GoTo ErrorHandler
   Call SubClearFields
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "Petty Cash Verification"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyV
            If BtnVerified.Enabled Then BtnVerified_Click
            KeyCode = 0
      End Select
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnVerified_Click()
   On Error GoTo ErrorHandler
   CN.BeginTrans
   Set Rs = New ADODB.Recordset
   sSql = "Select * FROM PettyCashHeader where ID = " & Val(TxtID.Text)
   Rs.Open sSql, CN, adOpenStatic, adLockOptimistic
   If Rs.RecordCount = 1 Then
      Rs!isVerify = 1
      Rs.Update
      MsgBox "Your Verification of Petty Cash is Sucessful", vbOKOnly, "Alert"
   End If
   CN.CommitTrans
   Unload Me
   Exit Sub
ErrorHandler:
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub

Private Property Get FormStatus() As FormMode
   On Error GoTo ErrorHandler
   'Nothing
   FormStatus = vMode
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
   'Based upon the value of vNewValue, we shall decide what controls to enable/disable
   On Error GoTo ErrorHandler
   vMode = vNewValue
   Select Case vNewValue
   Case Is = NewMode
      Call SubClearFields
      TxtID.Text = CN.Execute("select ID from PettyCashHeader where ToUserNo = " & vUser & " and EntryDate in (select max(EntryDate) from PettyCashHeader where ToUserNo = " & vUser & " and isverify = 0)").Fields(0).Value
      LoadGrid
   Case Is = OpenMode
   Case Is = changeMode
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo ErrorHandler
   Set FrmPettyCashVerification = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then
            ctl.Text = ""
         End If
      End If
   Next
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
