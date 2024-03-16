VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form DefCustomProductsandMeasurements 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "DefCPandMeasurements.frx":0000
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
      Left            =   11625
      TabIndex        =   5
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
         Height          =   1230
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Tag             =   "NC"
         Text            =   "DefCPandMeasurements.frx":0ECA
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
      Left            =   2003
      TabIndex        =   0
      Top             =   2490
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
      Left            =   7343
      TabIndex        =   1
      Top             =   2490
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
      stylesets(0).Picture=   "DefCPandMeasurements.frx":0F0E
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
      stylesets(1).Picture=   "DefCPandMeasurements.frx":16FC
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
      stylesets(2).Picture=   "DefCPandMeasurements.frx":2016
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
      Columns.Count   =   4
      Columns(0).Width=   2064
      Columns(0).Caption=   "ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5662
      Columns(1).Caption=   "Name"
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
      Columns(3).Width=   2011
      Columns(3).Caption=   "Unit"
      Columns(3).Name =   "Unit"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   10213
      _ExtentY        =   12144
      _StockProps     =   79
      Caption         =   "Custom Product Measurements"
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
      Left            =   11633
      TabIndex        =   3
      Top             =   9480
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
      MICON           =   "DefCPandMeasurements.frx":2032
      BC              =   14737632
      FC              =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9593
      Top             =   1230
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
            Picture         =   "DefCPandMeasurements.frx":204E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DefCPandMeasurements.frx":25F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin JeweledBut.JeweledButton BtnRefresh 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   10328
      TabIndex        =   2
      Top             =   9480
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
      MICON           =   "DefCPandMeasurements.frx":2FAD
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
      Left            =   12923
      TabIndex        =   8
      Top             =   1590
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Products and Measurements"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   4
      Top             =   270
      Width           =   4845
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Menu MnuCustomProducts 
      Caption         =   "Custom Products"
      Visible         =   0   'False
      Begin VB.Menu mniCreateNewCustomProduct 
         Caption         =   "Create new Custom Product"
      End
      Begin VB.Menu mniModifyPropertiesForCustomProduct 
         Caption         =   "Modify the Properties for this Custom Product"
      End
      Begin VB.Menu mniDeleteCustomProduct 
         Caption         =   "Delete this Custom Product"
      End
      Begin VB.Menu MniSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mniCreatenewMeasurement 
         Caption         =   "Create new Measurement for this Custom Product"
      End
      Begin VB.Menu mniModifyMeasurement 
         Caption         =   "Modify the Selected Measurement"
      End
      Begin VB.Menu mniDeleteMeasurement 
         Caption         =   "Delete this Measurement"
      End
   End
End
Attribute VB_Name = "DefCustomProductsandMeasurements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsPM As ADODB.Recordset
Dim vid As String

Private Sub BtnClose_Click()
    Unload Me
End Sub

Private Sub BtnRefresh_Click()
   RefreshMeasurementList
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
  SetWindowText Me.hWnd, "Orderd Product and Measurments"
  ShowPicture Me, 2
  AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
  HelpLocation Me
  Call RefreshMeasurementList
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

Private Sub RefreshMeasurementList()
  On Error GoTo ErrorHandler
  Dim X As Node
  TvwGroups.Nodes.Clear
  With CN.Execute("Select * FROM CustomProductsMeasurements where isdetailed=0 Order by ID")
    Do Until .EOF
      If IsNull(!ParentID) Then
        Set X = TvwGroups.Nodes.Add(, , "Account:" & !id, !Name, 1, 2)
        X.Tag = Abs(!isdetailed) & Abs(!IsLocked) & Abs(!iseditable)
      Else
        Set X = TvwGroups.Nodes.Add("Account:" & !ParentID, tvwChild, "Account:" & !id, !Name, 1, 2)
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
    If RsPM.Sort = "ID Asc" Then
      RsPM.Sort = "ID Desc"
      Grid.Columns(0).HeadStyleSet = "DESC"
      Grid.Columns(1).HeadStyleSet = ""
    Else
      RsPM.Sort = "ID Asc"
      Grid.Columns(0).HeadStyleSet = "ASC"
      Grid.Columns(1).HeadStyleSet = ""
    End If
  Else
    If RsPM.Sort = "Name Asc" Then
      RsPM.Sort = "Name Desc"
      Grid.Columns(1).HeadStyleSet = "DESC"
      Grid.Columns(0).HeadStyleSet = ""
    Else
      RsPM.Sort = "Name Asc"
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
    mniDeleteMeasurement.Enabled = CBool(Mid(Grid.Columns("Flags").Text, 3, 1)) And (Not CBool(Mid(Grid.Columns("Flags").Text, 2, 1)))
    mniCreatenewMeasurement.Enabled = Not CBool(Mid(Grid.Columns("Flags").Text, 2, 1))
    mniCreateNewCustomProduct.Enabled = False
    mniDeleteCustomProduct.Enabled = False
    mniModifyMeasurement.Enabled = CBool(Mid(Grid.Columns("Flags").Text, 3, 1)) And (Not CBool(Mid(Grid.Columns("Flags").Text, 2, 1)))
    mniModifyPropertiesForCustomProduct.Enabled = False
    Me.PopupMenu MnuCustomProducts
  End If
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub mniCreatenewMeasurement_Click()
  On Error GoTo ErrorHandler
  With DefCreateCustomProductMeasurement
    .ParaInID = ""
    .ParaInIsGroup = False
    .ParaInIsLocked = False
    .ParaInIsNew = True
    .ParaInParentName = CN.Execute("Select Name from CustomProductsMeasurements where ID = '" & Left(Replace(TvwGroups.SelectedItem.Key, "Account:", ""), 2) & "'").Fields(0)
    .ParaInParentID = Replace(TvwGroups.SelectedItem.Key, "Account:", "")
    .Show vbModal, Me
    If .ParaOutUpdateSuccess Then RefreshMeasurementList
  End With
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub mniCreateNewCustomProduct_Click()
  On Error GoTo ErrorHandler
  With DefCreateCustomProductMeasurement
    .ParaInID = ""
    .ParaInIsGroup = True
    .ParaInIsLocked = False
    .ParaInIsNew = True
    .ParaInParentName = CN.Execute("Select Name from CustomProductsMeasurements where ID = '" & Left(Replace(TvwGroups.SelectedItem.Key, "Account:", ""), 2) & "'").Fields(0)
    .ParaInParentID = IIf(Len(Replace(TvwGroups.SelectedItem.Key, "Account:", "")) = 1, "", Replace(TvwGroups.SelectedItem.Key, "Account:", ""))
    .Show vbModal, Me
    If .ParaOutUpdateSuccess Then RefreshMeasurementList
  End With
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub mniDeleteMeasurement_Click()
  On Error GoTo ErrorHandler
  Dim vtbl As String
  If MsgBox("Do you want to remove this Measurement?", vbQuestion + vbYesNo, "Alert") = vbNo Then Exit Sub
  vtbl = Common.ChildDataExists("CustomProductsMeasurements", "ID='" & Grid.Columns("ID").Text & "'", "")
  If vtbl <> "" Then
    MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
    Exit Sub
  End If
  'Call ActivityLog("Chart of Accounts", eDelete, , , Grid.Columns("ID").Text)
  CN.Execute ("Delete From CustomProductsMeasurements Where ID = '" & Grid.Columns("ID").Text & "'")
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub mniDeleteCustomProduct_Click()
  On Error GoTo ErrorHandler
  Dim vtbl As String
  If MsgBox("Do you want to remove this Custom Product?", vbQuestion + vbYesNo, "Alert") = vbNo Then Exit Sub
  vtbl = Common.ChildDataExists("CustomProductsMeasurements", "ID='" & Replace(TvwGroups.SelectedItem.Key, "Account:", "") & "'", "")
  If CN.Execute("Select count(*) FROM CustomProductsMeasurements Where ParentID = '" & Replace(TvwGroups.SelectedItem.Key, "Account:", "") & "'").Fields(0) > 0 Then
    MsgBox "You cannot remove this Custom Product because some child records exist for this Custom Product", vbCritical, "Error"
    Exit Sub
  ElseIf vtbl <> "" Then
    MsgBox "The record cannot be deleted because it exists in table : " & vtbl, vbCritical, "Error"
    Exit Sub
  End If
  'Call ActivityLog("Chart of Accounts", eDelete, , , Replace(TvwGroups.SelectedItem.Key, "Account:", ""))
  CN.Execute ("Delete From CustomProductsMeasurements Where ID = '" & Replace(TvwGroups.SelectedItem.Key, "Account:", "") & "'")
  TvwGroups.Nodes.Remove TvwGroups.SelectedItem.Index
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub mniModifyMeasurement_Click()
  On Error GoTo ErrorHandler
  
  With DefCreateCustomProductMeasurement
    .ParaInID = Grid.Columns("ID").Text
    .ParaInIsGroup = False
    .ParaInIsLocked = CN.Execute("Select islocked from CustomProductsMeasurements where ID = '" & Grid.Columns("ID").Text & "'").Fields(0)
    .ParaInIsNew = False
    .ParaInParentName = CN.Execute("Select Name from CustomProductsMeasurements where ID = '" & Left(Replace(TvwGroups.SelectedItem.Key, "Account:", ""), 1) & "'").Fields(0)
    .ParaInParentID = Replace(TvwGroups.SelectedItem.Key, "Account:", "")
    
    .Show vbModal, Me
    If .ParaOutUpdateSuccess Then RefreshMeasurementList
  End With
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub mniModifyPropertiesForCustomProduct_Click()
  On Error GoTo ErrorHandler
  
  With DefCreateCustomProductMeasurement
    .ParaInID = Replace(TvwGroups.SelectedItem.Key, "Account:", "")
    .ParaInIsGroup = True
    .ParaInIsLocked = CN.Execute("Select islocked from CustomProductsMeasurements where ID = '" & Replace(TvwGroups.SelectedItem.Key, "Account:", "") & "'").Fields(0)
    .ParaInIsNew = False
    If TypeName(TvwGroups.SelectedItem.Parent) <> "Nothing" Then .ParaInParentName = CN.Execute("Select Name from CustomProductsMeasurements where ID = '" & Left(Replace(TvwGroups.SelectedItem.Key, "Account:", ""), 1) & "'").Fields(0)
    If TypeName(TvwGroups.SelectedItem.Parent) <> "Nothing" Then .ParaInParentID = Replace(TvwGroups.SelectedItem.Parent.Key, "Account:", "")
    
    .Show vbModal, Me
    If .ParaOutUpdateSuccess Then RefreshMeasurementList
  End With
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub TvwGroups_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo ErrorHandler
   TvwGroups.SetFocus
   If Button = 2 And Shift = 0 Then
      mniDeleteMeasurement.Enabled = False
      mniCreatenewMeasurement.Enabled = (Len(vid) = 2)
      mniCreateNewCustomProduct.Enabled = (Len(vid) = 1)
      mniDeleteCustomProduct.Enabled = CBool(Mid(TvwGroups.SelectedItem.Tag, 3, 1))
      mniModifyMeasurement.Enabled = False
      mniModifyPropertiesForCustomProduct.Enabled = CBool(Mid(TvwGroups.SelectedItem.Tag, 3, 1))
      Me.PopupMenu MnuCustomProducts
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TvwGroups_NodeClick(ByVal Node As MSComctlLib.Node)
  On Error GoTo ErrorHandler
  
  vid = Mid(Node.Key, InStr(1, Node.Key, ":") + 1)
  Set RsPM = CN.Execute("SElect p.*, UnitName, cast(IsDetailed as varchar)+cast(IsLocked as varchar)+cast(IsEditable as varchar) as Flags FROM CustomProductsMeasurements p left outer join units u on u.unitid = p.unitid Where ParentID = '" & vid & "' AND IsDetailed=1")
  Set Grid.DataSource = RsPM
  Grid.Columns("ID").DataField = "ID"
  Grid.Columns("Name").DataField = "Name"
  Grid.Columns("Flags").DataField = "Flags"
  Grid.Columns("Unit").DataField = "UnitName"
  Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub
