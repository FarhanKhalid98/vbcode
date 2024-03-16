VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form DefChartOfAccounts 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "ChartOfAccounts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
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
      Height          =   1590
      Left            =   13320
      TabIndex        =   5
      Top             =   480
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
         Height          =   1230
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Tag             =   "NC"
         Text            =   "ChartOfAccounts.frx":0ECA
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
         TabIndex        =   7
         Top             =   90
         Width           =   135
      End
   End
   Begin MSComctlLib.TreeView TvwGroups 
      Height          =   6885
      Left            =   2108
      TabIndex        =   0
      Top             =   1860
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   12144
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6885
      Left            =   7463
      TabIndex        =   1
      Top             =   1860
      Width           =   5790
      ScrollBars      =   2
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   3
      stylesets(0).Name=   "DESC"
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
      stylesets(0).Picture=   "ChartOfAccounts.frx":0F0E
      stylesets(0).AlignmentPicture=   1
      stylesets(0).PictureMetaWidth=   353
      stylesets(0).PictureMetaHeight=   353
      stylesets(1).Name=   "ASC"
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
      stylesets(1).Picture=   "ChartOfAccounts.frx":16FC
      stylesets(1).AlignmentPicture=   1
      stylesets(1).PictureMetaWidth=   353
      stylesets(1).PictureMetaHeight=   353
      stylesets(2).Name=   "SelectedRow"
      stylesets(2).ForeColor=   16777215
      stylesets(2).BackColor=   8388608
      stylesets(2).HasFont=   -1  'True
      BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(2).Picture=   "ChartOfAccounts.frx":2016
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
      Columns.Count   =   3
      Columns(0).Width=   2064
      Columns(0).Caption=   "A/c #"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   7646
      Columns(1).Caption=   "A/c Name"
      Columns(1).Name =   "Name"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "Flags"
      Columns(2).Name =   "Flags"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   10213
      _ExtentY        =   12144
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
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   11738
      TabIndex        =   3
      Top             =   8850
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
      MICON           =   "ChartOfAccounts.frx":2032
      BC              =   14737632
      FC              =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7965
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartOfAccounts.frx":204E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartOfAccounts.frx":25F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin JeweledBut.JeweledButton BtnRefresh 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10433
      TabIndex        =   2
      Top             =   8850
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Refresh"
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
      MICON           =   "ChartOfAccounts.frx":2FAD
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
      Left            =   11295
      TabIndex        =   8
      Top             =   495
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chart of Accounts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
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
      Width           =   2490
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Menu MnuAccounts 
      Caption         =   "Accounts"
      Visible         =   0   'False
      Begin VB.Menu mniCreateNewChildGroup 
         Caption         =   "Create new Child Group"
      End
      Begin VB.Menu mniModifyPropertiesForGroup 
         Caption         =   "Modify the Properties for this Group"
      End
      Begin VB.Menu mniDeleteGroup 
         Caption         =   "Delete this Group"
      End
      Begin VB.Menu MniSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mniCreatenewAccounts 
         Caption         =   "Create new Accounts for this Group"
      End
      Begin VB.Menu mniModifyAccount 
         Caption         =   "Modify the Selected Account"
      End
      Begin VB.Menu mniDeleteAccount 
         Caption         =   "Delete this Account"
      End
   End
End
Attribute VB_Name = "DefChartOfAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsAccounts As ADODB.Recordset

Private Sub BtnClose_Click()
    Unload Me
End Sub

Private Sub BtnRefresh_Click()
   RefreshAccountList
End Sub

Private Sub Form_Activate()
   'incomplete
End Sub

Private Sub LblClose_Click()
   FraHelp.Visible = False
End Sub

Private Sub LblHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LblHelp.ForeColor = &H800000
   FraHelp.ZOrder 0
   FraHelp.Visible = True
End Sub

Private Sub LblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If LblHelp.FontUnderline = True Then Exit Sub
   LblHelp.FontUnderline = True
End Sub

Private Sub LblHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LblHelp.ForeColor = vbWhite
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
  SetWindowText Me.hWnd, "Chart of Accounts"
  ShowPicture Me, 2
  AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
  HelpLocation Me
  Call RefreshAccountList
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      keybd_event 9, 1, 1, 1
      KeyCode = 0
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      KeyCode = 0
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyH
               FraHelp.ZOrder 0
               FraHelp.Visible = True
               KeyCode = 0
         Case vbKeyR
            If BtnRefresh.Enabled Then BtnRefresh_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   End If
End Sub

Private Sub RefreshAccountList()
  On Error GoTo ErrorHandler
  Dim X As Node
  TvwGroups.Nodes.Clear
  With cn.Execute("Select * FROM Chartofaccounts where isdetailed=0 Order by AccountNo")
    Do Until .EOF
      If IsNull(!ParentAccountNo) Then
        Set X = TvwGroups.Nodes.Add(, , "Account:" & !AccountNo, !AccountName, 1, 2)
        X.Tag = Abs(!isdetailed) & Abs(!IsLocked) & Abs(!iseditable)
      Else
        Set X = TvwGroups.Nodes.Add("Account:" & !ParentAccountNo, tvwChild, "Account:" & !AccountNo, !AccountName, 1, 2)
        X.Tag = Abs(!isdetailed) & Abs(!IsLocked) & Abs(!iseditable)
      End If
      .MoveNext
    Loop
  End With
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
  On Error GoTo ErrorHandler
  If ColIndex = 0 Then
    If RsAccounts.Sort = "AccountNo Asc" Then
      RsAccounts.Sort = "AccountNo Desc"
      Grid.Columns(0).HeadStyleSet = "DESC"
      Grid.Columns(1).HeadStyleSet = ""
    Else
      RsAccounts.Sort = "AccountNo Asc"
      Grid.Columns(0).HeadStyleSet = "ASC"
      Grid.Columns(1).HeadStyleSet = ""
    End If
  Else
    If RsAccounts.Sort = "AccountName Asc" Then
      RsAccounts.Sort = "AccountName Desc"
      Grid.Columns(1).HeadStyleSet = "DESC"
      Grid.Columns(0).HeadStyleSet = ""
    Else
      RsAccounts.Sort = "AccountName Asc"
      Grid.Columns(1).HeadStyleSet = "ASC"
      Grid.Columns(0).HeadStyleSet = ""
    End If
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  If Button = 2 And Shift = 0 Then
    mniDeleteAccount.Enabled = CBool(Mid(Grid.Columns("Flags").Text, 3, 1)) And (Not CBool(Mid(Grid.Columns("Flags").Text, 2, 1)))
    If ((TvwGroups.SelectedItem.Key Like "*61") Or (TvwGroups.SelectedItem.Key Like "*62") Or (TvwGroups.SelectedItem.Key Like "*63")) Then
      mniCreatenewAccounts.Enabled = False
    Else
      mniCreatenewAccounts.Enabled = Not CBool(Mid(Grid.Columns("Flags").Text, 2, 1))
    End If
    mniCreateNewChildGroup.Enabled = False
    mniDeleteGroup.Enabled = False
    mniModifyAccount.Enabled = CBool(Mid(Grid.Columns("Flags").Text, 3, 1)) And (Not CBool(Mid(Grid.Columns("Flags").Text, 2, 1)))
    mniModifyPropertiesForGroup.Enabled = False
    Me.PopupMenu MnuAccounts
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub mniCreatenewAccounts_Click()
  On Error GoTo ErrorHandler
  With DefCreateAccounts
    .ParaInAccountNo = ""
    .ParaInIsGroup = False
    .ParaInIsLocked = False
    .ParaInIsNew = True
    .ParaInParentAccountName = cn.Execute("Select AccountType from chartofaccounts where accountno = " & Left(Replace(TvwGroups.SelectedItem.Key, "Account:", ""), 1)).Fields(0)
    .ParaInParentAccountNo = Replace(TvwGroups.SelectedItem.Key, "Account:", "")
    .Show vbModal, Me
    If .ParaOutUpdateSuccess Then RefreshAccountList
  End With
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub mniCreateNewChildGroup_Click()
  On Error GoTo ErrorHandler
  With DefCreateAccounts
    .ParaInAccountNo = ""
    .ParaInIsGroup = True
    .ParaInIsLocked = False
    .ParaInIsNew = True
    .ParaInParentAccountName = cn.Execute("Select Accountname from chartofaccounts where accountno = " & Left(Replace(TvwGroups.SelectedItem.Key, "Account:", ""), 1)).Fields(0)
    .ParaInParentAccountNo = Replace(TvwGroups.SelectedItem.Key, "Account:", "")
    .Show vbModal, Me
    If .ParaOutUpdateSuccess Then RefreshAccountList
  End With
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub mniDeleteAccount_Click()
  On Error GoTo ErrorHandler
  Dim vtbl As String
  If MsgBox("Do you want to remove this Account?", vbQuestion + vbYesNo, "Alert") = vbNo Then Exit Sub
  vtbl = Common.ChildDataExists("ChartofAccounts", "AccountNo = " & Grid.Columns("ID").Text, "")
  If vtbl <> "" Then
    MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
    Exit Sub
  End If
  Call ActivityLog("Chart of Accounts", eDelete, , , Grid.Columns("ID").Text)
  cn.Execute ("Delete From ChartOfAccounts Where AccountNo = " & Val(Grid.Columns("ID").Text))
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub mniDeleteGroup_Click()
  On Error GoTo ErrorHandler
  Dim vtbl As String
  If MsgBox("Do you want to remove this group?", vbQuestion + vbYesNo, "Alert") = vbNo Then Exit Sub
  vtbl = Common.ChildDataExists("ChartofAccounts", "AccountNo = " & Replace(TvwGroups.SelectedItem.Key, "Account:", ""), "")
  If cn.Execute("Select count(*) FROM ChartOfAccounts Where ParentAccountNo = '" & Replace(TvwGroups.SelectedItem.Key, "Account:", "") & "'").Fields(0) > 0 Then
    MsgBox "You cannot remove this group because some child records exist for this group", vbCritical, "Error"
    Exit Sub
  ElseIf vtbl <> "" Then
    MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
    Exit Sub
  End If
  Call ActivityLog("Chart of Accounts", eDelete, , , Replace(TvwGroups.SelectedItem.Key, "Account:", ""))
  cn.Execute ("Delete From ChartOfAccounts Where AccountNo = " & Replace(TvwGroups.SelectedItem.Key, "Account:", ""))
  TvwGroups.Nodes.Remove TvwGroups.SelectedItem.Index
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub mniModifyAccount_Click()
  On Error GoTo ErrorHandler
  
  With DefCreateAccounts
    .ParaInAccountNo = Grid.Columns("ID").Text
    .ParaInIsGroup = False
    .ParaInIsLocked = cn.Execute("Select islocked from chartofaccounts where accountno = " & Val(Grid.Columns("ID").Text)).Fields(0)
    .ParaInIsNew = False
    .ParaInParentAccountName = cn.Execute("Select Accountname from chartofaccounts where accountno = " & Left(Replace(TvwGroups.SelectedItem.Key, "Account:", ""), 1)).Fields(0)
    .ParaInParentAccountNo = Replace(TvwGroups.SelectedItem.Key, "Account:", "")
    
    .Show vbModal, Me
    If .ParaOutUpdateSuccess Then RefreshAccountList
  End With
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub mniModifyPropertiesForGroup_Click()
  On Error GoTo ErrorHandler
  
  With DefCreateAccounts
    .ParaInAccountNo = Replace(TvwGroups.SelectedItem.Key, "Account:", "")
    .ParaInIsGroup = True
    .ParaInIsLocked = cn.Execute("Select islocked from chartofaccounts where accountno = " & Replace(TvwGroups.SelectedItem.Key, "Account:", "")).Fields(0)
    .ParaInIsNew = False
    If TypeName(TvwGroups.SelectedItem.Parent) <> "Nothing" Then .ParaInParentAccountName = cn.Execute("Select Accountname from chartofaccounts where accountno = " & Left(Replace(TvwGroups.SelectedItem.Key, "Account:", ""), 1)).Fields(0)
    If TypeName(TvwGroups.SelectedItem.Parent) <> "Nothing" Then .ParaInParentAccountNo = Replace(TvwGroups.SelectedItem.Parent.Key, "Account:", "")
    
    .Show vbModal, Me
    If .ParaOutUpdateSuccess Then RefreshAccountList
  End With
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub TvwGroups_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  TvwGroups.SetFocus
  If Button = 2 And Shift = 0 Then
    mniDeleteAccount.Enabled = False
    mniCreatenewAccounts.Enabled = Not ((Replace(TvwGroups.SelectedItem.Key, "Account:", "") Like "61") Or (Replace(TvwGroups.SelectedItem.Key, "Account:", "") Like "62") Or (Replace(TvwGroups.SelectedItem.Key, "Account:", "") Like "63"))
    mniCreateNewChildGroup.Enabled = True
    mniDeleteGroup.Enabled = CBool(Mid(TvwGroups.SelectedItem.Tag, 3, 1))
    mniModifyAccount.Enabled = False
    mniModifyPropertiesForGroup.Enabled = CBool(Mid(TvwGroups.SelectedItem.Tag, 3, 1))
    
    Me.PopupMenu MnuAccounts
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub TvwGroups_NodeClick(ByVal Node As MSComctlLib.Node)
  On Error GoTo ErrorHandler
  Dim vAccNO As String
  vAccNO = Mid(Node.Key, InStr(1, Node.Key, ":") + 1)
  Set RsAccounts = cn.Execute("SElect *,cast(IsDetailed as varchar)+cast(IsLocked as varchar)+cast(IsEditable as varchar) as Flags FROM ChartOfAccounts Where ParentAccountNo = '" & vAccNO & "' AND IsDetailed=1")
  Set Grid.DataSource = RsAccounts
  Grid.Columns("ID").DataField = "AccountNo"
  Grid.Columns("Name").DataField = "AccountName"
  Grid.Columns("Flags").DataField = "Flags"
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub
