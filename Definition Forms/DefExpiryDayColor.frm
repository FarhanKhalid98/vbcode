VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form DefExpiryDayColor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15450
   Icon            =   "DefExpiryDayColor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   770
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1030
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbExpiryColor 
      Height          =   315
      ItemData        =   "DefExpiryDayColor.frx":0ECA
      Left            =   8670
      List            =   "DefExpiryDayColor.frx":0EDA
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   5610
      Width           =   1740
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
      Left            =   11115
      TabIndex        =   21
      Top             =   720
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
         TabIndex        =   22
         Tag             =   "NC"
         Text            =   "DefExpiryDayColor.frx":0EFB
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
         TabIndex        =   23
         Top             =   90
         Width           =   135
      End
   End
   Begin VB.TextBox TxtFilter 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2145
      MaxLength       =   30
      TabIndex        =   8
      Top             =   2888
      Width           =   4545
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4650
      Left            =   2145
      TabIndex        =   9
      Top             =   3218
      Width           =   4815
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
      stylesets(0).Picture=   "DefExpiryDayColor.frx":0F86
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
      Columns(0).Width=   2487
      Columns(0).Caption=   "ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4948
      Columns(1).Caption=   "Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   8493
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
   Begin JeweledBut.JeweledButton BtnNew 
      Height          =   420
      Left            =   2505
      TabIndex        =   10
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
      MICON           =   "DefExpiryDayColor.frx":0FA2
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOpen 
      Height          =   420
      Left            =   3825
      TabIndex        =   11
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
      MICON           =   "DefExpiryDayColor.frx":0FBE
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   5145
      TabIndex        =   12
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
      MICON           =   "DefExpiryDayColor.frx":0FDA
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   8295
      TabIndex        =   5
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
      MICON           =   "DefExpiryDayColor.frx":0FF6
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9615
      TabIndex        =   6
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
      MICON           =   "DefExpiryDayColor.frx":1012
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10935
      TabIndex        =   7
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
      MICON           =   "DefExpiryDayColor.frx":102E
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtName 
      Height          =   315
      Left            =   5730
      TabIndex        =   3
      Top             =   705
      Visible         =   0   'False
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   50
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
   Begin SITextBox.Txt TxtID 
      Height          =   315
      Left            =   8580
      TabIndex        =   0
      Top             =   3698
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   2
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
   Begin JeweledBut.JeweledButton BtnAccount 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6510
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "nc"
      Top             =   1695
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
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
      MICON           =   "DefExpiryDayColor.frx":104A
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtAccountNo 
      Height          =   315
      Left            =   5730
      TabIndex        =   4
      Top             =   1695
      Visible         =   0   'False
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   11
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
      IntegralPoint   =   10
   End
   Begin SITextBox.Txt TxtAccountName 
      Height          =   315
      Left            =   6870
      TabIndex        =   18
      Tag             =   "nc"
      Top             =   1695
      Visible         =   0   'False
      Width           =   3585
      _ExtentX        =   6324
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
   Begin SITextBox.Txt TxtDayFrom 
      Height          =   315
      Left            =   8580
      TabIndex        =   1
      Top             =   4800
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      IntegralPoint   =   3
   End
   Begin SITextBox.Txt TxtDayTo 
      Height          =   315
      Left            =   9600
      TabIndex        =   2
      Top             =   4800
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   2
      IntegralPoint   =   3
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Remaining Days From Expiry and select Color"
      Height          =   195
      Left            =   8550
      TabIndex        =   29
      Top             =   4230
      Width           =   3615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8670
      TabIndex        =   28
      Top             =   5400
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Days To"
      Height          =   195
      Left            =   9630
      TabIndex        =   26
      Top             =   4560
      Width           =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Days From"
      Height          =   195
      Left            =   8580
      TabIndex        =   25
      Top             =   4560
      Width           =   750
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
      TabIndex        =   24
      Top             =   450
      Width           =   435
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C No"
      Height          =   195
      Left            =   5730
      TabIndex        =   20
      Top             =   1485
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Name"
      Height          =   195
      Left            =   6810
      TabIndex        =   19
      Top             =   1485
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Day Color"
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
      TabIndex        =   16
      Top             =   270
      Width           =   2220
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   504
      X2              =   503
      Y1              =   186.533
      Y2              =   528.533
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Day Color"
      Height          =   195
      Left            =   2145
      TabIndex        =   15
      Top             =   2685
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   195
      Left            =   8580
      TabIndex        =   14
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   5730
      TabIndex        =   13
      Top             =   465
      Visible         =   0   'False
      Width           =   420
   End
End
Attribute VB_Name = "DefExpiryDayColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim vMode As FormMode
Dim vIsNewRecord As Boolean 'will flag whether the record is new or existing one.

Private Sub BtnClear_Click()
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('ExpiryDayColor'" & "," & TxtID.Text & ",Null,'Cleared','" & Date & "','" & Time & "',6,'Cleared'," & vUser & ")")
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
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtAccountNo.Name: If FunSelectAccount(ssFunctionKey, True) = True Then CmbExpiryColor.SetFocus
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
    cn.Execute ("Insert Into UserActivities values ('ExpiryDayColor'" & "," & TxtID.Text & ",Null,'Closed','" & Date & "','" & Time & "',7,'Closed'," & vUser & ")")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Unload Me
End Sub

Private Sub BtnDelete_Click()
  On Error GoTo ErrorHandler
  
   ''''''''''''' User Authentication ''''''''''''''
   vUserAction = UserAuthentication("MniExpiryDayColor", vUser, ObjUserSecurity.IsAdministrator, eUserDelete)
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
  
  Dim vtbl As String
  If Rs.RecordCount > 0 Then
    If MsgBox("Do you really want to remove this record?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then Exit Sub
    vtbl = Common.ChildDataExists("ExpiryDayColor", "ExpiryColorID='" & Rs!ExpiryColorID & "'", "")
    If vtbl <> "" Then
      MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
      Exit Sub
    End If
    Call ActivityLog("ExpiryDayColor", eDelete, TxtID.Text)
    '''''''''''''''''''''''''''''''''''''User Activities'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cn.Execute ("Insert Into UserActivities values ('ExpiryDayColor'" & "," & TxtID.Text & ",Null,'Removed','" & Date & "','" & Time & "',3,'Removed'," & vUser & ")")
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
   vUserAction = UserAuthentication("MniExpiryDayColor", vUser, ObjUserSecurity.IsAdministrator, IIf(vIsNewRecord = True, eUserNewRecord, eUserEdit))
   If vUserAction <> "" Then
      MsgBox vUserAction, vbCritical, "Error"
      Exit Sub
   End If
   ''''''''''''' '''''''''''''''''''' ''''''''''''''
   
   If vIsNewRecord = False Then Call ActivityLog("ExpiryDayColor", eEdit, TxtID.Text)
   
   Call UserActivities
   
   Rs.Filter = "ExpiryColorID = " & TxtID.Text
   If vIsNewRecord Then
      Rs.AddNew
      Rs!ExpiryColorID = TxtID.Text
   End If
   Rs!ExpiryColorName = CmbExpiryColor.Text & " From " & Val(TxtDayFrom.Text) & " To " & Val(TxtDayTo.Text) & " Days."
   Rs!DayFrom = Val(TxtDayFrom.Text)
   Rs!DayTo = Val(TxtDayTo.Text)
   Rs!ExpiryColor = CmbExpiryColor.Text
   Rs.Update
   Rs.Filter = ""
   If vIsNewRecord = True Then Call ActivityLog("ExpiryDayColor", eAdd, TxtID.Text)
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunValidation() As Boolean
  On Error GoTo ErrorHandler
  If vIsNewRecord Then
    If Trim(TxtID.Text) = "" Then
      MsgBox "Please specify a ExpiryDayColor ID", vbExclamation, "Alert"
      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
      Exit Function
    End If
    If Trim(TxtDayFrom.Text) = "" Then
      MsgBox "Please specify a Day From", vbExclamation, "Alert"
      If TxtDayFrom.Enabled And TxtDayFrom.Visible Then TxtDayFrom.SetFocus
      Exit Function
    End If
    If Trim(TxtDayTo.Text) = "" Then
      MsgBox "Please specify a Day To", vbExclamation, "Alert"
      If TxtDayTo.Enabled And TxtDayTo.Visible Then TxtDayTo.SetFocus
      Exit Function
    End If
'    If Len(Trim(TxtID.Text)) < 2 Then
'      MsgBox "The ExpiryDayColor ID must be two characters long", vbExclamation, "Alert"
'      If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
'      Exit Function
'    End If
    If cn.Execute("Select count(*) from ExpiryDayColor where ExpiryColorID = " & TxtID.Text).Fields(0) > 0 Then
        MsgBox "This ExpiryDayColor ID already exists. The ExpiryDayColor ID must be unique", vbExclamation, "Alert"
        If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
        Exit Function
    End If
'    Select Case Asc(UCase(Left(TxtID.Text, 1)))
'      Case 65 To 90
'      Case 48 To 57
'      Case Else
'        MsgBox "The Group ID must contain numeric/alphabetical characters only", vbExclamation, "Alert"
'        If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
'        Exit Function
'    End Select
'    Select Case Asc(UCase(Right(TxtID.Text, 1)))
'      Case 65 To 90
'      Case 48 To 57
'      Case Else
'        MsgBox "The Group ID must contain numeric/alphabetical characters only", vbExclamation, "Alert"
'        If TxtID.Enabled And TxtID.Visible Then TxtID.SetFocus
'        Exit Function
'    End Select
  End If
'  If Trim(TxtName.Text) = "" Then
'    MsgBox "Please specify a ExpiryDayColor name", vbExclamation, "Alert"
'    If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
'    Exit Function
'  End If

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
   SetWindowText Me.hWnd, "ExpiryDayColors"
   HelpLocation Me
'   CmbExpiryColor.ListIndex = 0
   Set Rs = New ADODB.Recordset
   Rs.Open "Select * FROM ExpiryDayColor", cn, adOpenDynamic, adLockOptimistic
   Set Grid.DataSource = Rs
   Grid.Columns("ID").DataField = "ExpiryColorID"
   Grid.Columns("Name").DataField = "ExpiryColorName"
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
         TxtName.Text = ""
         TxtID.Text = ""
         TxtAccountNo.Text = ""
         TxtAccountName.Text = ""
         TxtDayFrom.Text = ""
         TxtDayTo.Text = ""
         CmbExpiryColor.ListIndex = 0
         TxtFilter.Enabled = False
         TxtFilter.BackColor = &HE0E0E0
         TxtID.Text = FunGetMaxID()
         TxtID.Enabled = True
         TxtName.Enabled = True
         BtnNew.Enabled = False
         BtnOpen.Enabled = False
         BtnDelete.Enabled = False
         BtnSave.Enabled = False
         BtnClear.Enabled = True
         TxtName.Enabled = True
         TxtID.Enabled = False
         BtnAccount.Enabled = True
         TxtAccountNo.Enabled = True
         CmbExpiryColor.Enabled = True
         TxtDayFrom.Enabled = True
         TxtDayTo.Enabled = True
         Grid.Enabled = False
'         If TxtName.Visible And TxtName.Enabled Then TxtName.SetFocus
         If TxtDayFrom.Visible And TxtDayFrom.Enabled Then TxtDayFrom.SetFocus
         TxtFilter.Text = ""
         vIsNewRecord = True
     Case Is = OpenMode
         TxtFilter.Enabled = False
         TxtFilter.BackColor = &HE0E0E0
         TxtName.Enabled = True
         TxtID.Enabled = False
         CmbExpiryColor.Enabled = True
         TxtDayFrom.Enabled = True
         TxtDayTo.Enabled = True
         BtnAccount.Enabled = True
         TxtAccountNo.Enabled = True
         If TxtName.Visible = True Then TxtName.SetFocus Else TxtDayFrom.SetFocus
         TxtFilter.Text = ""
         BtnNew.Enabled = False
         BtnOpen.Enabled = False
         BtnDelete.Enabled = False
         BtnClear.Enabled = True
         Grid.Enabled = False
         vIsNewRecord = False
     Case Is = ChangeMode
         BtnSave.Enabled = True
     Case Is = SelectionMode
         Grid.Enabled = True
         TxtFilter.Enabled = True
         TxtFilter.BackColor = vbWhite
         BtnNew.Enabled = True
         BtnOpen.Enabled = True
         BtnDelete.Enabled = True
         BtnSave.Enabled = False
         BtnClear.Enabled = False
         TxtName.Enabled = False
         TxtID.Enabled = False
         TxtAccountNo.Enabled = False
         BtnAccount.Enabled = False
         CmbExpiryColor.Enabled = False
         TxtDayFrom.Enabled = False
         TxtDayTo.Enabled = False
         Call Grid_RowColChange(0, 0)
         Grid.SetFocus
         TxtFilter.Text = ""
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
      Set DefCommissionDiscRange = Nothing
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
      TxtDayFrom.Text = Rs!DayFrom
      TxtDayTo.Text = Rs!DayTo
      CmbExpiryColor.Text = Rs!ExpiryColor
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
   Rs.Find "ExpiryColorName like '" & Replace(TxtFilter.Text, "'", "''") & "%'", , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Function FunGetMaxID() As String
  On Error GoTo ErrorHandler
  FunGetMaxID = cn.Execute("Select isnull(max(ExpiryColorID),0) + 1 from ExpiryDayColor").Fields(0)
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Private Sub BtnAccount_Click()
   If FunSelectAccount(ssButton, False) = True Then
      If BtnSave.Enabled Then BtnSave.SetFocus
   Else
      If TxtAccountNo.Enabled Then TxtAccountNo.SetFocus
   End If
End Sub

Private Function FunSelectAccount(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim vStrSQL As String
    If CallerName = ssButton Or CallerName = ssFunctionKey Then
        SchAccounts.ParaInAllowListSelection = True
        SchAccounts.ParaInWhereClause = ""
        SchAccounts.ParaInDetail = ""
        SchAccounts.Show vbModal, Me
        If SchAccounts.ParaOutAccountNo = "" Then FunSelectAccount = False: Exit Function
        TxtAccountNo.Text = SchAccounts.ParaOutAccountNo
    End If
    '---------------------------
    vStrSQL = " Select * FROM ChartOfAccounts where AccountNo='" & TxtAccountNo.Text & "'"
    With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
          TxtAccountName.Text = !AccountName
          FunSelectAccount = True
          .Close
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
          Exit Function
      Else
          FunSelectAccount = False
          .Close
          TxtAccountNo.Text = ""
          TxtAccountName.Text = ""
          If BtnSave.Enabled = False Then FormStatus = ChangeMode
      End If
   End With
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub TxtAccountNo_Change()
   If TxtAccountNo.Visible = False Then Exit Sub
   If ActiveControl.Name <> TxtAccountNo.Name Then Exit Sub
   If TxtAccountName.Text <> "" Then TxtAccountName.Text = ""
End Sub

Private Sub TxtAccountNo_Validate(Cancel As Boolean)
   If Me.ActiveControl.Name <> TxtAccountNo.Name Then Exit Sub
   On Error GoTo ErrorHandler
   If TxtAccountName.Text <> "" Then Exit Sub
   Dim vTemp As Boolean
   vTemp = Not FunSelectAccount(ssValidate, True)
   If vTemp = True Then
      vTemp = Not FunSelectAccount(ssButton, False)
   End If
   Cancel = vTemp
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub UserActivities()
    If vIsNewRecord = False Then
    With cn.Execute("Select  * from ExpiryDayColor where ExpiryColorID =" & TxtID.Text)
        If TxtName.Text <> !ExpiryColorName Then
            cn.Execute ("Insert Into UserActivities values ('ExpiryDayColor'" & "," & TxtID.Text & ", Null , 'Updated ExpiryDayColor Name-" & !ExpiryColorName & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
'        If TxtAccountNo.Text <> !AccountNo Then
'            cn.Execute ("Insert Into UserActivities values ('ExpiryDayColor'" & "," & TxtID.Text & ", Null , 'Updated Account No-" & !AccountNo & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
'        End If
        If CmbExpiryColor.Text <> !ExpiryColor Then
            cn.Execute ("Insert Into UserActivities values ('ExpiryDayColor'" & "," & TxtID.Text & ", Null , 'Updated ExpiryColor-" & !ExpiryColor & "','" & Date & "','" & Time & "',2,'Updated'," & vUser & ")")
        End If
    End With
   Else
        cn.Execute ("Insert Into UserActivities values ('ExpiryDayColor'" & "," & TxtID.Text & ", Null ,'Saved','" & Date & "','" & Time & "',1,'Saved'," & vUser & ")")
   End If
End Sub
