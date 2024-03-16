VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmNumberList 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8663
      TabIndex        =   9
      Top             =   2100
      Width           =   1230
   End
   Begin VB.TextBox TxtNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1823
      MaxLength       =   15
      TabIndex        =   5
      Text            =   "+92"
      Top             =   2415
      Width           =   2250
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7538
      TabIndex        =   0
      Top             =   9885
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
      MICON           =   "FrmNumberList.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnMembers 
      Height          =   645
      Left            =   6983
      TabIndex        =   2
      Top             =   3645
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1138
      TX              =   "Add Members"
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
      MICON           =   "FrmNumberList.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6945
      Left            =   1830
      TabIndex        =   3
      Tag             =   "C"
      ToolTipText     =   "Product Entry"
      Top             =   2820
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   12250
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   6233
      TabIndex        =   4
      Top             =   9885
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
      MICON           =   "FrmNumberList.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnOK 
      Height          =   420
      Left            =   4928
      TabIndex        =   6
      Top             =   9885
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "OK"
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
      MICON           =   "FrmNumberList.frx":0054
      BC              =   14737632
      FC              =   0
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   7350
      Left            =   8663
      TabIndex        =   7
      Tag             =   "C"
      ToolTipText     =   "Product Entry"
      Top             =   2415
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   12965
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin JeweledBut.JeweledButton BtnTransfer 
      Height          =   420
      Left            =   6983
      TabIndex        =   8
      Top             =   6375
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "<<"
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
      MICON           =   "FrmNumberList.frx":0070
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnTextFile 
      Height          =   645
      Left            =   6983
      TabIndex        =   10
      Top             =   2730
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1138
      TX              =   "Add From Text File"
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
      MICON           =   "FrmNumberList.frx":008C
      BC              =   14737632
      FC              =   0
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   13073
      Top             =   9840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin JeweledBut.JeweledButton BtnParties 
      Height          =   645
      Left            =   6983
      TabIndex        =   11
      Top             =   4560
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1138
      TX              =   "Add Parties"
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
      MICON           =   "FrmNumberList.frx":00A8
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnCounterNo 
      Height          =   645
      Left            =   6983
      TabIndex        =   12
      Top             =   5430
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1138
      TX              =   "Add Counter"
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
      MICON           =   "FrmNumberList.frx":00C4
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number List"
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
      TabIndex        =   1
      Top             =   270
      Width           =   1695
   End
   Begin VB.Image ImgExit 
      Height          =   360
      Left            =   13268
      Top             =   1590
      Width           =   330
   End
End
Attribute VB_Name = "FrmNumberList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Item As ListItem
Dim vSQL As String


Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
   If ListView1.ListItems.Count = 0 Then Exit Sub
   ListView1.ListItems.Clear
   ListView2.ListItems.Clear
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnCounterNo_Click()
   On Error GoTo ErrorHandler
   ListView2.ListItems.Clear
   vSQL = "Select  distinct('+92' + Right(Replace(customername, '-', ''), 10)) Mobile from saleheader where isnumeric(Right(Replace(customername, '-', ''), 10)) = 1"
   With CN.Execute(vSQL)
      While Not .EOF
         If IsNumeric(!Mobile) And Len(!Mobile) = 13 Then
            Set Item = ListView2.ListItems.Add(, , "621")
            Item.SubItems(1) = !Mobile
            .MoveNext
         Else
            .MoveNext
         End If
      Wend
   End With
   PopulateListView
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnMembers_Click()
   On Error GoTo ErrorHandler
   ListView2.ListItems.Clear
   With CN.Execute("select MemberID,  '+92' + right(replace(mobile,'-',''),10) as Mobile from Members where len(right(replace(mobile,'-',''),10)) = 10")
      While Not .EOF
         If IsNumeric(!Mobile) And Len(!Mobile) = 13 Then
            Set Item = ListView2.ListItems.Add(, , !MemberID)
            Item.SubItems(1) = !Mobile
'            .MoveNext
         End If
         .MoveNext
      Wend
   End With
   PopulateListView
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnOK_Click()
   On Error GoTo ErrorHandler
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open "Select * FROM SMSList", CN, adOpenDynamic, adLockPessimistic
   For i = 1 To ListView1.ListItems.Count
      Rs.AddNew
      Rs!SNo = i
      Rs!MobileNo = ListView1.ListItems(i).Text
      Rs!Message = FrmSMS.strMessage
      Rs.Update
'      vSQL = "insert into SMSList values('" & i & "','" & ListView1.ListItems(i).Text & "','" & FrmSMS.TextBox1.Text & "')"
'      CN.Execute vSQL
   Next i
   FrmSMS.PopulateListView
   Rs.Close
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnTextFile_Click()
   On Error GoTo ErrorHandler
   CD1.FileName = ""
   CD1.DialogTitle = "For Text File List"
   CD1.InitDir = App.Path
   CD1.Filter = "(Text Files)|*.txt"
   CD1.ShowOpen
   Dim vID As Integer
   vID = 0
   ListView2.ListItems.Clear
   If CD1.FileName <> "" Then
      Dim vString As String
      Open CD1.FileName For Input As #1
      Do Until EOF(1)
         vID = vID + 1
         Line Input #1, vString
         vString = Replace(vString, ",", "")
         vString = Replace(vString, "-", "")
         If Len(vString) = 11 And Left(vString, 1) = "0" Then
            vString = "+92" & Right(vString, 10)
         End If
         Set Item = ListView2.ListItems.Add(, , vID)
         Item.SubItems(1) = vString
      Loop
      Close #1
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnTransfer_Click()
   
Open App.Path & "\Config.ini" For Input As #1
   Line Input #1, vConnStr
'   Do Until EOF(1)
'      Line Input #1, vString
'      Debug.Print vString
'   Loop
   Close #1
   
   On Error GoTo ErrorHandler
   For i = 1 To ListView2.ListItems.Count
      If ListView2.ListItems(i).Checked = True Then
         Set Item = ListView1.ListItems.Add(, , ListView2.ListItems(i).ListSubItems(1).Text)
      End If
   Next i
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ChkAll_Click()
   On Error GoTo ErrorHandler
   For i = 1 To ListView2.ListItems.Count
      ListView2.ListItems(i).Checked = Abs(ChkAll.Value)
   Next i
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If UCase(ActiveControl.Name) = UCase(TxtNumber.Name) Then
         Set Item = ListView1.ListItems.Add(, , TxtNumber.Text)
         TxtNumber.Text = "+92"
         TxtNumber.SelStart = Len(TxtNumber.Text)
      Else
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub PopulateListView()
   On Error GoTo ErrorHandler
   ListView1.ListItems.Clear
'   With CN.Execute("select * from SMSList")
'      While Not .EOF
'         Set Item = ListView1.ListItems.Add(, , !SNo)
'         Item.SubItems(1) = !MobileNo
'         Item.SubItems(2) = !Message
'         .MoveNext
'      Wend
'   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "Number List"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   

   ListView1.FullRowSelect = True
   ListView1.ListItems.Clear
   ListView1.ColumnHeaders.Add , , "Numers", 150, 0
   ListView1.View = lvwReport
   
   ListView2.FullRowSelect = True
   ListView2.ListItems.Clear
   ListView2.ColumnHeaders.Add , , "Sr No", 100, 0
   ListView2.ColumnHeaders.Add , , "Mobile No", 150, 0
   ListView2.View = lvwReport
'   PopulateListView

   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub BtnParties_Click()
   On Error GoTo ErrorHandler
   ListView2.ListItems.Clear
   With CN.Execute("select PartyID,  '+92' + right(replace(mobile,'-',''),10) as Mobile from parties where len(right(replace(mobile,'-',''),10)) = 10 and mobile is not null and mobile <> '' and PartyID like '62%' ")
      While Not .EOF
         Set Item = ListView2.ListItems.Add(, , !PartyID)
         Item.SubItems(1) = !Mobile
'         .MoveNext
      Wend
      .MoveNext
   End With
   PopulateListView
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage

End Sub
