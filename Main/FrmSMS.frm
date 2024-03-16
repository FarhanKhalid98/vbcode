VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmSMS 
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
   Begin VB.CheckBox ChkUrdu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Urdu Text"
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
      Left            =   9405
      TabIndex        =   9
      Top             =   4020
      Width           =   1230
   End
   Begin VB.Timer Timer1 
      Left            =   12285
      Top             =   2850
   End
   Begin VB.TextBox TxtDelayInSec 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10800
      TabIndex        =   4
      Text            =   "6"
      Top             =   3885
      Width           =   510
   End
   Begin VB.TextBox TxtMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   9360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4650
      Width           =   4110
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   11205
      TabIndex        =   0
      Top             =   7575
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
      MICON           =   "FrmSMS.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSend 
      Height          =   420
      Left            =   11790
      TabIndex        =   3
      Top             =   3840
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Send"
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
      MICON           =   "FrmSMS.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7305
      Left            =   1800
      TabIndex        =   6
      Tag             =   "C"
      ToolTipText     =   "Product Entry"
      Top             =   2850
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   12885
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
      Left            =   9900
      TabIndex        =   7
      Top             =   7575
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
      MICON           =   "FrmSMS.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   2580
      Left            =   9360
      TabIndex        =   8
      ToolTipText     =   "Textbox1"
      Top             =   4515
      Width           =   4110
      VariousPropertyBits=   752896027
      ForeColor       =   0
      BorderStyle     =   1
      Size            =   "7250;4551"
      SpecialEffect   =   0
      FontName        =   "@Arial Unicode MS"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delay in (Sec)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10440
      TabIndex        =   5
      Top             =   3525
      Width           =   1260
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SMS"
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
      Width           =   660
   End
   Begin VB.Image ImgExit 
      Height          =   360
      Left            =   13290
      Top             =   1755
      Width           =   330
   End
End
Attribute VB_Name = "FrmSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Item As ListItem
Dim vSQL As String
Dim ModeValue As Boolean
Dim UniCode As Variant
Public strMessage As String

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
   If ListView1.ListItems.Count = 0 Then Exit Sub
   vSQL = "Delete From SMSList"
   CN.Execute vSQL
   PopulateListView
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnSend_Click()
   On Error GoTo ErrorHandler
   If TxtMessage.Visible = True Then
      strMessage = TxtMessage.Text
   Else
      strMessage = TextBox1.Text
   End If
   FrmNumberList.Show
'   With CN.Execute("select MemberID,  '+92' + right(replace(mobile,'-',''),10) as Mobile from Members where len(right(replace(mobile,'-',''),10)) = 10")
'      While Not .EOF
'         vSQL = "insert into SMSList values('" & !MemberID & "','" & !Mobile & "','" & TxtMessage.Text & "')"
'         CN.Execute vSQL
'         .MoveNext
'      Wend
'   End With
'   PopulateListView
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ChkUrdu_Click()

If ChkUrdu.Value = 1 Then
   TxtMessage.Visible = False
   TextBox1.Visible = True
End If
If ChkUrdu.Value = 0 Then
   TxtMessage.Visible = True
   TextBox1.Visible = False
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
     keybd_event 9, 1, 1, 1
     KeyCode = 0
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Public Sub PopulateListView()
   On Error GoTo ErrorHandler
   ListView1.ListItems.Clear
   If Rs.State = adStateOpen Then Rs.Close
'   With CN.Execute("select * from SMSList")
   Rs.Open "Select * FROM SMSList", CN, adOpenDynamic, adLockPessimistic
      While Not Rs.EOF
         Set Item = ListView1.ListItems.Add(, , Rs!SNo)
         Item.SubItems(1) = Rs!MobileNo
         Item.SubItems(2) = Rs!Message
         Rs.MoveNext
      Wend
'   End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "SMS Sender"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   Timer1.Interval = Val(TxtDelayInSec.Text) * 1000
   

   ListView1.FullRowSelect = True
   ListView1.ListItems.Clear
   ListView1.ColumnHeaders.Add , , "SNo", 50, 0
   ListView1.ColumnHeaders.Add , , "MobileNo", 120, 0
   ListView1.ColumnHeaders.Add , , "Message", 400, 0
   ListView1.View = lvwReport

   PopulateListView
   
   TextBox1.Visible = False

   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub Timer1_Timer()
   On Error GoTo ErrorHandler
   If ListView1.ListItems.Count = 0 Then Exit Sub
   vSQL = "insert into MessageOut(MessageTo, MessageFrom, MessageText, MessageType) select MobileNo,'',Message,'' from SMSList where SNO = '" & ListView1.SelectedItem.Text & "'"
   CN.Execute vSQL
   vSQL = "Delete From SMSList where SNO = '" & ListView1.SelectedItem.Text & "'"
   CN.Execute vSQL
   PopulateListView
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtDelayInSec_Change()
   On Error GoTo ErrorHandler
   Timer1.Interval = Val(TxtDelayInSec.Text) * 1000
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub TextBox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
        ''''''''''''''''''''''''''''''''''''''''''''''''
        ''                                             ''
        '' There are the KeyDown-Event behaviours for  ''
        '' Enter, Space, Delete & Tab keys to set      ''
        '' Behavior in Textbox1.Text, keys will behave ''
        '' as Normal Text writing behavior.            ''
        ''                                             ''
        '''''''''''''''''''''''''''''''''''''''''''''''''

   If ModeValue = False Then
      'Space Key Behavior
         If KeyCode = 32 Then
         UniCode = &H20
         TextBox1.Text = TextBox1.Text + ChrW(UniCode)
         KeyCode = 0

        'Enter Key Behavior
'        ElseIf KeyCode = 13 Then
'        UniCode = &HA
'        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
'        KeyCode = 0

        'Horizontal Tab Behavior
'        ElseIf KeyCode = 9 Then
'        UniCode = &H9
'        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
'        KeyCode = 0

         'Delete Key Behavior
         ElseIf KeyCode = 127 Then
         UniCode = &H7F
         TextBox1.Text = TextBox1.Text + ChrW(UniCode)
         KeyCode = 0
         End If
   End If
        
        'This Function Got End There
End Sub

Private Sub TextBox1_KeyPress(KeyAscii As MSForms.ReturnInteger)

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
        If KeyAscii = 97 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H627
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'b Key Behavior
        ElseIf KeyAscii = 98 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H628
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'c Key Behavior
        ElseIf KeyAscii = 99 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H686
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'd Key Behavior
        ElseIf KeyAscii = 100 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'e Key Behavior
        ElseIf KeyAscii = 101 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H639
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'f Key Behavior
        ElseIf KeyAscii = 102 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H641
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'g Key Behavior
        ElseIf KeyAscii = 103 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6AF
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'h Key Behavior
        ElseIf KeyAscii = 104 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6BE
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'i Key Behavior
        ElseIf KeyAscii = 105 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6CC
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'j Key Behavior
        ElseIf KeyAscii = 106 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'k Key Behavior
        ElseIf KeyAscii = 107 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6A9
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'l Key Behavior
        ElseIf KeyAscii = 108 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H644
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'm Key Behavior
        ElseIf KeyAscii = 109 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H645
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'n Key Behavior
        ElseIf KeyAscii = 110 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H646
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'o Key Behavior
        ElseIf KeyAscii = 111 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6C1
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'p Key Behavior
        ElseIf KeyAscii = 112 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H67E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'q Key Behavior
        ElseIf KeyAscii = 113 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H642
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'r Key Behavior
        ElseIf KeyAscii = 114 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H631
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        's Key Behavior
        ElseIf KeyAscii = 115 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H633
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        't Key Behavior
        ElseIf KeyAscii = 116 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'u Key Behavior
        ElseIf KeyAscii = 117 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H621
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'v Key Behavior
        ElseIf KeyAscii = 118 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H637
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'w Key Behavior
        ElseIf KeyAscii = 119 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H648
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'x Key Behavior
        ElseIf KeyAscii = 120 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H634
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'y Key Behavior
        ElseIf KeyAscii = 121 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6D2
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'z Key Behavior
        ElseIf KeyAscii = 122 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H632
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)


        ' For Capital Latter's Behaviors

        'A Key Behavior
        ElseIf KeyAscii = 65 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H622
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'B Key Behavior
        ElseIf KeyAscii = 66 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFBB0
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'C Key Behavior
        ElseIf KeyAscii = 67 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'D Key Behavior
        ElseIf KeyAscii = 68 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H688
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'E Key Behavior
        ElseIf KeyAscii = 69 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H650
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'F Key Behavior
        ElseIf KeyAscii = 70 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H652
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'G Key Behavior
        ElseIf KeyAscii = 71 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H63A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'H Key Behavior
        ElseIf KeyAscii = 72 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'I Key Behavior
        ElseIf KeyAscii = 73 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H649
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'J Key Behavior
        ElseIf KeyAscii = 74 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H636
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'K Key Behavior
        ElseIf KeyAscii = 75 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'L Key Behavior
        ElseIf KeyAscii = 76 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFEFB
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'M Key Behavior
        ElseIf KeyAscii = 77 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H66B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'N Key Behavior
        ElseIf KeyAscii = 78 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6BA
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'O Key Behavior
        ElseIf KeyAscii = 79 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H629
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'P Key Behavior
        ElseIf KeyAscii = 80 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H64F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'Q Key Behavior
        ElseIf KeyAscii = 81 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H626
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'R Key Behavior
        ElseIf KeyAscii = 82 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H691
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'S Key Behavior
        ElseIf KeyAscii = 83 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H635
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'T Key Behavior
        ElseIf KeyAscii = 84 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H679
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'U Key Behavior
        ElseIf KeyAscii = 85 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H626
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'V Key Behavior
        ElseIf KeyAscii = 86 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H638
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'W Key Behavior
        ElseIf KeyAscii = 87 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H624
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'Z Key Behavior
        ElseIf KeyAscii = 88 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H698
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'Y Key Behavior
        ElseIf KeyAscii = 89 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFBAF
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        'Z Key Behavior
        ElseIf KeyAscii = 90 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H630
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)


        'For Numaric Key's Behaviors

        '0 Key Behavior
        ElseIf KeyAscii = 48 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H660
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '1 Key Behavior
        ElseIf KeyAscii = 49 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H661
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '2 Key Behavior
        ElseIf KeyAscii = 50 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H662
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '3 Key Behavior
        ElseIf KeyAscii = 51 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H663
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '4 Key Behavior
        ElseIf KeyAscii = 52 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H664
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '5 Key Behavior
        ElseIf KeyAscii = 53 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H665
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '6 Key Behavior
        ElseIf KeyAscii = 54 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H666
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '7 Key Behavior
        ElseIf KeyAscii = 55 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H667
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '8 Key Behavior
        ElseIf KeyAscii = 56 Or TextBox1.SelText <> "" Then
        UniCode = &H668
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '9 Key Behavior
        ElseIf KeyAscii = 57 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H669
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        ' Numaric Keys with 'Shift' Behavior

        ') Key Behavior
        ElseIf KeyAscii = 41 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFD3F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '! Key Behavior
        ElseIf KeyAscii = 33 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H21
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '@ Key Behavior
        ElseIf KeyAscii = 64 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H40
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '# Key Behavior
        ElseIf KeyAscii = 35 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H23
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '$ Key Behavior
        ElseIf KeyAscii = 36 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H24
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '% Key Behavior
        ElseIf KeyAscii = 37 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H66A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '^ Key Behavior
        ElseIf KeyAscii = 94 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '& Key Behavior
        ElseIf KeyAscii = 38 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H26
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '* Key Behavior
        ElseIf KeyAscii = 42 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H66D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '( Key Behavior
        ElseIf KeyAscii = 40 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFD3E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)


        'For Special Characters

        'Symbols

        '? Key Behavior
        ElseIf KeyAscii = 63 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H61F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '/ Key Behavior
        ElseIf KeyAscii = 47 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        ', Key Behavior
        ElseIf KeyAscii = 44 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H60C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '. Key Behavior
        ElseIf KeyAscii = 46 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H640
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '_ Key Behavior
        ElseIf KeyAscii = 95 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '- Key Behavior
        ElseIf KeyAscii = 45 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '+ Key Behavior
        ElseIf KeyAscii = 43 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '= Key Behavior
        ElseIf KeyAscii = 61 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H3D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        ': Key Behavior
        ElseIf KeyAscii = 58 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H3A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '; Key Behavior
        ElseIf KeyAscii = 59 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H201C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '< Key Behavior
        ElseIf KeyAscii = 60 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H64E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '> Key Behavior
        ElseIf KeyAscii = 62 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H650
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '{ Key Behavior
        ElseIf KeyAscii = 123 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2018
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '} Key Behavior
        ElseIf KeyAscii = 125 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2019
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '[ Key Behavior
        ElseIf KeyAscii = 91 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '] Key Behavior
        ElseIf KeyAscii = 93 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '| Key Behavior
        ElseIf KeyAscii = 124 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H7C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '\ Key Behavior
        ElseIf KeyAscii = 92 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '~ Key Behavior
        ElseIf KeyAscii = 126 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H64B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '` Key Behavior
        ElseIf KeyAscii = 96 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H64D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '" Key Behavior
        ElseIf KeyAscii = 34 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2190
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        '' Key Behavior
        ElseIf KeyAscii = 39 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H201D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)

        End If
        KeyAscii = 0
  End If

        'This Function Got End There

End Sub


