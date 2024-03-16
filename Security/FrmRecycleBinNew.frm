VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmRecycleBinNew 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDetail 
      Caption         =   "Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5865
      Left            =   10125
      TabIndex        =   13
      Top             =   2655
      Width           =   5145
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridDetail 
         Height          =   5580
         Left            =   90
         TabIndex        =   14
         Top             =   270
         Width           =   4980
         ScrollBars      =   2
         _Version        =   196616
         RecordSelectors =   0   'False
         stylesets.count =   1
         stylesets(0).Name=   "Select"
         stylesets(0).ForeColor=   0
         stylesets(0).BackColor=   13817275
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
         stylesets(0).Picture=   "FrmRecycleBinNew.frx":0000
         AllowUpdate     =   0   'False
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
         SelectTypeRow   =   1
         RowNavigation   =   1
         ForeColorEven   =   0
         BackColorEven   =   15724527
         BackColorOdd    =   16777215
         RowHeight       =   423
         ActiveRowStyleSet=   "Select"
         Columns.Count   =   5
         Columns(0).Width=   1640
         Columns(0).Caption=   "ID"
         Columns(0).Name =   "ID"
         Columns(0).CaptionAlignment=   2
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "FormNo"
         Columns(1).Name =   "FormNo"
         Columns(1).DataField=   "Column 9"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   4286
         Columns(2).Caption=   "Description"
         Columns(2).Name =   "Description"
         Columns(2).DataField=   "Column 3"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   2328
         Columns(3).Caption=   "Amount"
         Columns(3).Name =   "Amount"
         Columns(3).DataField=   "Column 10"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
         Columns(4).Caption=   "SerialNo"
         Columns(4).Name =   "SerialNo"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   8784
         _ExtentY        =   9843
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
   End
   Begin VB.ComboBox CmbAction 
      Height          =   315
      ItemData        =   "FrmRecycleBinNew.frx":001C
      Left            =   11010
      List            =   "FrmRecycleBinNew.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1395
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.ComboBox CmbUser 
      Height          =   315
      Left            =   3255
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2160
      Width           =   2790
   End
   Begin VB.ComboBox CmbFilter 
      Height          =   315
      Left            =   6090
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2160
      Width           =   3150
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7808
      TabIndex        =   6
      Top             =   9495
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
      MICON           =   "FrmRecycleBinNew.frx":0020
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnRestore 
      Height          =   420
      Left            =   6300
      TabIndex        =   5
      Top             =   9495
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "ReStore"
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
      MICON           =   "FrmRecycleBinNew.frx":003C
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFromDate 
      Height          =   330
      Left            =   9270
      TabIndex        =   2
      Top             =   2160
      Width           =   1395
      _Version        =   65543
      _ExtentX        =   2461
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "/"
      Format          =   "dd/MM/yyyy"
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpToDate 
      Height          =   330
      Left            =   10710
      TabIndex        =   3
      Top             =   2160
      Width           =   1395
      _Version        =   65543
      _ExtentX        =   2461
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "/"
      Format          =   "dd/MM/yyyy"
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6000
      Left            =   45
      TabIndex        =   12
      Top             =   2610
      Width           =   10050
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   11
      stylesets.count =   1
      stylesets(0).Name=   "Select"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   12566463
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
      stylesets(0).Picture=   "FrmRecycleBinNew.frx":0058
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      MultiLine       =   0   'False
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
      SelectTypeRow   =   1
      RowNavigation   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   11
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "BinID"
      Columns(0).Name =   "BinID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3387
      Columns(1).Caption=   "Bin Date"
      Columns(1).Name =   "BinDate"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   7
      Columns(1).NumberFormat=   "dd//mm/yy h:mm:ss"
      Columns(1).FieldLen=   256
      Columns(2).Width=   1270
      Columns(2).Caption=   "User No"
      Columns(2).Name =   "UserNo"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2381
      Columns(3).Caption=   "User Name"
      Columns(3).Name =   "UserName"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3704
      Columns(4).Caption=   "FormType"
      Columns(4).Name =   "FormType"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1561
      Columns(5).Caption=   "ID"
      Columns(5).Name =   "ID"
      Columns(5).CaptionAlignment=   0
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "SID"
      Columns(6).Name =   "SID"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   2037
      Columns(7).Caption=   "Date"
      Columns(7).Name =   "Date"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   7
      Columns(7).NumberFormat=   "dd/mm/yy"
      Columns(7).FieldLen=   256
      Columns(8).Width=   1879
      Columns(8).Caption=   "Amount"
      Columns(8).Name =   "Amount"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   926
      Columns(9).Caption=   "Sel"
      Columns(9).Name =   "Selection"
      Columns(9).Alignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   11
      Columns(9).FieldLen=   256
      Columns(9).Style=   2
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "FormNo"
      Columns(10).Name=   "FormNo"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   17727
      _ExtentY        =   10583
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Actions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   11010
      TabIndex        =   11
      Top             =   1170
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3255
      TabIndex        =   10
      Top             =   1935
      Width           =   900
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Form Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6090
      TabIndex        =   9
      Top             =   1935
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "-------From Date To Date ---------"
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
      Left            =   9360
      TabIndex        =   8
      Top             =   1935
      Width           =   2655
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recycle Bin"
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
      TabIndex        =   7
      Top             =   270
      Width           =   1680
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "FrmRecycleBinNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim RsBody As New ADODB.Recordset
Dim vStrSQL, vStrPara, vActionNo, vSSID, vRSID As String
Dim vCounter As Integer
Public ParaOutUserNo As String
Dim vOrder As String, vDirection As String, vCol As Byte, vFilter As String, vUser As String, vAction As String

Private Sub BtnClose_Click()
   On Error GoTo ErrorHandler
   Me.ParaOutUserNo = ""
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnRestore_Click()
On Error GoTo ErrorHandler
      If MsgBox("Do you want to Restore Data?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
      Call ReStoreData
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbAction_Click()
   On Error GoTo ErrorHandler
   If CmbAction.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbAction.Name Then Exit Sub
   Select Case CmbAction.ListIndex
   Case 0
      vAction = ""
   Case 1
      vAction = " and h.Remarks = 'Delete'"
   Case 2
      vAction = " and h.Remarks = 'Clear'"
   End Select
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbFilter_Click()
   On Error GoTo ErrorHandler
   If CmbFilter.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbFilter.Name Then Exit Sub
   If CmbFilter.ListIndex = 0 Then
      vFilter = ""
   Else
      vFilter = " and Bin.FormNo = " & CmbFilter.ItemData(CmbFilter.ListIndex)
   End If
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbUser_Click()
   On Error GoTo ErrorHandler
   If CmbUser.Visible = False Then Exit Sub
   If ActiveControl.Name <> CmbUser.Name Then Exit Sub
   If CmbUser.ListIndex = 0 Then
      vUser = ""
   Else
      vUser = " and u.UserNo =" & CmbUser.ItemData(CmbUser.ListIndex)
   End If
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub DtpFromDate_Change()
   On Error GoTo ErrorHandler
   If DtpFromDate.IsDateValid = False Then Exit Sub
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub DtpToDate_Change()
   On Error GoTo ErrorHandler
   If DtpToDate.IsDateValid = False Then Exit Sub
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyEscape Then Call BtnClose_Click
'   If KeyCode = vbKeyReturn Then
'      Select Case ActiveControl.Name
'      Case Grid.Name, TxtUserName.Name
'         Call BtnSelect_Click
'      End Select
'   End If
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hwnd, "Recycle Bin"
   
   CmbFilter.Clear
   CmbFilter.AddItem "-- All Types --"
   With CN.Execute("Select * FROM " & vBinDataBase & ".dbo.FormBin Order by FormNo ASC")
      Do Until .EOF
          CmbFilter.AddItem !FormType
          CmbFilter.ItemData(CmbFilter.NewIndex) = !FormNo
          .MoveNext
      Loop
   End With
   
   CmbUser.Clear
   CmbUser.AddItem "-- All Users --"
   With CN.Execute("Select * FROM Users")
      Do Until .EOF
          CmbUser.AddItem !UserName
          CmbUser.ItemData(CmbUser.NewIndex) = !UserNo
          .MoveNext
      Loop
   End With
   
   vActionNo = ObjRegistry.ByDefaultActionNo
   CmbAction.Clear
   CmbAction.AddItem "-- All Actions --"
   With CN.Execute("Select * FROM " & vBinDataBase & ".dbo.ActionBin " & IIf(vActionNo = "", "", " Where ActionNo in (" & vActionNo & " )"))
      Do Until .EOF
          CmbAction.AddItem !ActionName
          CmbAction.ItemData(CmbAction.NewIndex) = !ActionNo
          .MoveNext
      Loop
   End With
   
   DtpFromDate.DateValue = Date - 10
   DtpToDate.DateValue = Date
   CmbFilter.ListIndex = 0
   CmbUser.ListIndex = 0
   CmbAction.ListIndex = 0
   vOrder = " order by TrxDate Desc"
   LoadGrid
   Me.ParaOutUserNo = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   vStrSQL = "Select BinID, BinDate, Bin.FormNo, FormType, TrxSID, TrxID, TrxDate, ActionUserNo, UserName, TotalAmount From " & vbCrLf _
            & "( " & vbCrLf _
            & ""
   ''''''''''' SaleHeaderBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, H.SID TrxSID, H.billid TrxID, h.billdate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.SaleHeaderBin H " & vbCrLf _
              & " WHERE isReplace = 0 and h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
   
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' SaleReturnHeaderBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, H.SID TrxSID, H.Returnid TrxID, h.Returndate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.SaleReturnHeaderBin H " & vbCrLf _
              & " WHERE isReplace = 0 and h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
   
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' ReplacementHeaderBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, SID TrxSID, H.Replaceid TrxID, h.Replacedate TrxDate, ActionUserNo,  (SaleAmount - ReturnAmount) TotalAmount from " & vBinDataBase & ".dbo.ReplacementHeaderBin H " & vbCrLf _
              & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
   
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' SaleOrderHeaderBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.OrderID TrxID, h.Orderdate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.SaleOrderHeaderBin H " & vbCrLf _
              & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
   
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' PurchaseHeaderBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.Purid TrxID, h.PurchaseDate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.PurchaseHeaderBin H " & vbCrLf _
              & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
                 
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' PurchaseReturnHeaderBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.Returnid TrxID, h.Returndate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.PurchaseReturnHeaderBin H " & vbCrLf _
              & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
   
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' PurchaseOrderHeaderBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.Orderid TrxID, h.OrderDate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.PurchaseOrderHeaderBin H " & vbCrLf _
              & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
  
  vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' DebitVouchersBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.VoucherNo TrxID, h.VoucherDate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.DebitVouchersBin H " & vbCrLf _
            & " Inner join (Select VoucherNo, Sum(Amount) TotalAmount From " & vBinDataBase & ".dbo.DebitVouchersBodyBin Group by VoucherNo) B on B.VoucherNo= H.VoucherNo " & vbCrLf _
            & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
            
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' CreditVouchersBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.VoucherNo TrxID, h.VoucherDate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.CreditVouchersBin H " & vbCrLf _
            & " Inner join (Select VoucherNo, Sum(Amount) TotalAmount From " & vBinDataBase & ".dbo.CreditVouchersBodyBin Group by VoucherNo) B on B.VoucherNo= H.VoucherNo " & vbCrLf _
            & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
            
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' JournalVouchersBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.VoucherNo TrxID, h.VoucherDate TrxDate, ActionUserNo,  0 TotalAmount from " & vBinDataBase & ".dbo.JournalVouchersBin H " & vbCrLf _
            & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
            
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' AdvanceVouchersBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.VoucherNo TrxID, h.VoucherDate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.AdvanceVouchersBin H " & vbCrLf _
            & " Inner join (Select VoucherNo, Sum(Amount) TotalAmount From " & vBinDataBase & ".dbo.AdvanceVouchersBodyBin Group by VoucherNo) B on B.VoucherNo= H.VoucherNo " & vbCrLf _
            & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
     
            
    vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' RecoveryHeaderBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.RecoveryID TrxID, h.RecoveryDate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.RecoveryHeaderBin H " & vbCrLf _
            & " Inner join (Select RecoveryID, Sum(Amount) TotalAmount From " & vBinDataBase & ".dbo.RecoveryCustomerBin Group by RecoveryID) B on B.RecoveryID= H.RecoveryID " & vbCrLf _
            & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
            
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' RecoveryHeaderBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.RecoveryID TrxID, h.RecoveryDate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.RecoveryHeaderBin H " & vbCrLf _
            & " Inner join (Select RecoveryID, Sum(Amount) TotalAmount From " & vBinDataBase & ".dbo.RecoveryInvoiceBin Group by RecoveryID) B on B.RecoveryID= H.RecoveryID " & vbCrLf _
            & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
   
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' PaymentHeaderBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.PaymentID TrxID, h.PaymentDate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.PaymentHeaderBin H " & vbCrLf _
            & " Inner join (Select PaymentID, Sum(Amount) TotalAmount From " & vBinDataBase & ".dbo.PaymentInvoiceBin Group by PaymentID) B on B.PaymentID= H.PaymentID " & vbCrLf _
            & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"

   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' PaymentHeaderBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.PaymentID TrxID, h.PaymentDate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.PaymentHeaderBin H " & vbCrLf _
            & " Inner join (Select PaymentID, Sum(Amount) TotalAmount From " & vBinDataBase & ".dbo.PaymentVenderBin Group by PaymentID) B on B.PaymentID= H.PaymentID " & vbCrLf _
            & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
            
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' StockWastageHeaderBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.WastageID TrxID, h.WastageDate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.StockWastageHeaderBin H " & vbCrLf _
            & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
          
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' StockAdjustmentHeaderBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.AdjustmentID TrxID, h.AdjustmentDate TrxDate, ActionUserNo,  TotalAmount from " & vBinDataBase & ".dbo.StockAdjustmentHeaderBin H " & vbCrLf _
            & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
   
  vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' AdminClosingBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.ID TrxID, h.EntryDate TrxDate, ActionUserNo,  TotalSale TotalAmount from " & vBinDataBase & ".dbo.AdminClosingBin H " & vbCrLf _
            & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
   
   vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
   ''''''''''' SalarieBin '''''''''''
   vStrSQL = vStrSQL & " Select H.BinID, H.BinDate, H.FormNo, '' TrxSID, H.SalaryID TrxID, h.EntryDate TrxDate, ActionUserNo,  Salary TotalAmount from " & vBinDataBase & ".dbo.SalariesBin H " & vbCrLf _
            & " WHERE h.Bindate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'"
            
   vStrSQL = vStrSQL & "" & vbCrLf _
         & ")Bin " & vbCrLf _
            & ""
         
   vStrSQL = vStrSQL & "Inner Join Bin.dbo.FormBin FB on FB.FormNo= Bin.FormNo" & vbCrLf _
             & "Inner Join Users U on U.UserNO= Bin.ActionUserNo " & vbCrLf _
             & vAction & vFilter & vUser & vOrder & vDirection
             
   
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Grid.Redraw = False
   Grid.MoveFirst
   Grid.RemoveAll
   Grid.AllowAddNew = True
   While Not Rs.EOF
      Grid.Columns("BinID").Text = Rs!BinID
      Grid.Columns("BinDate").Text = Rs!BinDate
      Grid.Columns("FormNo").Value = Rs!FormNo
      Grid.Columns("FormType").Text = Rs!FormType
      Grid.Columns("SID").Value = Rs!TrxSID
      Grid.Columns("ID").Text = Rs!TrxID
      Grid.Columns("Date").Text = Rs!TrxDate
      Grid.Columns("UserNo").Text = Rs!ActionUserNo
      Grid.Columns("UserName").Text = Rs!UserName
      Grid.Columns("Amount").Value = Rs!TotalAmount
      Grid.Columns("Selection").Value = 0
      Grid.AddNew
      Rs.MoveNext
   Wend
   Grid.MoveFirst
   Grid.Redraw = True
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo ErrorHandler
   Select Case Grid.Columns(ColIndex).Name
      Case "UserName"
            vOrder = " order by U.UserName "
      Case "FormType"
            vOrder = " order by FormType "
      Case "ActionName"
            vOrder = " order by ActionName "
      Case "Selection"
            vOrder = " "
      Case "ID"
            vOrder = " "
      Case "Date"
            vOrder = " "
      Case "Amount"
            vOrder = " "
      Case Else
            vOrder = " order by Bin." & Grid.Columns(ColIndex).Name
   End Select
   
  If Grid.Columns(ColIndex).Name = "BinDate" Or Grid.Columns(ColIndex).Name = "UserNo" Or Grid.Columns(ColIndex).Name = "UserName" Or Grid.Columns(ColIndex).Name = "FormType" Then
      If vCol = ColIndex Then
         vDirection = IIf(vDirection = " Asc", " Desc", " Asc")
      Else
         vDirection = " Asc"
      End If
   Else
      vDirection = " "
   End If
   

   vCol = ColIndex
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo ErrorHandler
   Call LoadGridDetail
   Exit Sub
ErrorHandler:
   If Err.Description = "Overflow" Then
      Resume Next
      Exit Sub
  End If
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   On Error GoTo ErrorHandler
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadGridDetail()
On Error GoTo ErrorHandler
   If Grid.Rows <= 1 Then Exit Sub
   
   If Trim(Grid.Columns("FormNo").Value) = eFrmSaleInvoicePOS Or Trim(Grid.Columns("FormNo").Value) = eFrmSaleInvoiceDIS Then
   
      vStrSQL = " Select SerialNo, b.productId ID,  'Qty:- ' + cast (floor(isnull(qtypack * multiplier,0) + qty) as varchar) +  ' Disc% ' + cast (b.DiscPer as varchar)  Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.SaleBodyBin b " & vbCrLf _
                  & " WHERE b.sid = '" & Grid.Columns("SID").Text & "'"
                  
   ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmSaleReturnInvoicePOS Or Trim(Grid.Columns("FormNo").Value) = eFrmSaleReturnInvoiceDIS Then
      
      vStrSQL = " Select SerialNo, b.productId ID, 'Qty:- ' + cast (floor(isnull(qtypack * multiplier,0) + qty) as varchar) +  ' Disc% ' + cast (b.DiscPer as varchar)  Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.SaleReturnBodyBin b " & vbCrLf _
                  & " WHERE b.sid = '" & Grid.Columns("SID").Text & "'"
                  
   ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmReplacementInvoice Then
   
      vStrSQL = " Select SerialNo, b.productId ID, 'Qty:- ' + cast (floor(isnull(qtypack * multiplier,0) +qty) as varchar) +  ' Disc% ' + cast (b.DiscPer as varchar) + ' In '  Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.SaleReturnBodyBin b " & vbCrLf _
                  & " inner join " & vBinDataBase & ".dbo.ReplacementHeaderBin H on H.RSID = B.Sid " & vbCrLf _
                  & " WHERE H.sid = '" & Grid.Columns("SID").Text & "'"
                  
      vStrSQL = vStrSQL & "" & vbCrLf _
            & "Union ALL " & vbCrLf _
            & ""
            
      vStrSQL = vStrSQL & "Select SerialNo, b.productId ID,  'Qty:- ' + cast (floor(isnull(qtypack * multiplier,0) +qty) as varchar) +  ' Disc% ' + cast (b.DiscPer as varchar)  + ' Out '  Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.SaleBodyBin b " & vbCrLf _
                  & " inner join " & vBinDataBase & ".dbo.ReplacementHeaderBin H on H.SSID = B.Sid " & vbCrLf _
                  & " WHERE H.sid = '" & Grid.Columns("SID").Text & "'"
                  
   ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmSaleOrderPOS Or Trim(Grid.Columns("FormNo").Value) = eFrmSaleOrderDIS Then
   
      vStrSQL = " Select SerialNo, b.productId ID,  'Qty:- ' + cast (floor(isnull(qtypack * multiplier,0) + qty) as varchar) +  ' Disc% ' + cast (b.DiscPer as varchar)  Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.SaleOrderBodyBin b " & vbCrLf _
                  & " WHERE b.Orderid = '" & Grid.Columns("ID").Text & "' And OrderDate = '" & Grid.Columns("Date").Text & "'"
                  
   ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmPurchaseInvoice Then
      
      vStrSQL = " Select SerialNo, b.productId ID, 'Qty:- ' + cast (floor(isnull(qtypack * multiplier,0) + qtyloose) as varchar) +  ' Disc% ' + cast (b.DiscPer as varchar)  Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.PurchaseBodyBin b " & vbCrLf _
                  & " WHERE b.purid = '" & Grid.Columns("ID").Text & "' And PurchaseDate = '" & Grid.Columns("Date").Text & "'"
                     
   ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmPurchaseReturnInvoice Then
      
      vStrSQL = " Select SerialNo, b.productId ID, 'Qty:- ' + cast (floor(isnull(qtypack * multiplier,0) +qtyloose) as varchar) +  ' Disc% ' + cast (b.DiscPer as varchar)  Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.PurchaseReturnBodyBin b " & vbCrLf _
                  & " WHERE b.Returnid = '" & Grid.Columns("ID").Text & "' And ReturnDate = '" & Grid.Columns("Date").Text & "'"
                  
   ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmPurchaseOrder Then
      
      vStrSQL = " Select SerialNo, b.productId ID, 'Qty:- ' + cast (floor(isnull(qtypack * multiplier,0) + qtyloose) as varchar) +  ' Disc% ' + cast (b.DiscPer as varchar)  Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.PurchaseOrderBodyBin b " & vbCrLf _
                  & " WHERE b.Orderid = '" & Grid.Columns("ID").Text & "' And OrderDate = '" & Grid.Columns("Date").Text & "'"

    ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmCreditVoucher Then
      
      vStrSQL = " Select SerialNo, b.AccountNo ID, Narration Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.CreditVouchersBodyBin b " & vbCrLf _
                  & " WHERE b.VoucherNo = '" & Grid.Columns("ID").Text & "'"
                  
    ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmDebitVoucher Then
      
      vStrSQL = " Select SerialNo, b.AccountNo ID, Narration Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.DebitVouchersBodyBin b " & vbCrLf _
                  & " WHERE b.VoucherNo = '" & Grid.Columns("ID").Text & "'"
                  
    ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmJournalVoucher Then
      
      vStrSQL = " Select SerialNo, b.AccountNo ID, Narration Description, Abs(Credit - Debit) Amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.JournalVouchersBodyBin b " & vbCrLf _
                  & " WHERE b.VoucherNo = '" & Grid.Columns("ID").Text & "'"
                  
   ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmAdvanceVoucher Then
      
      vStrSQL = " Select SerialNo, b.AccountNo ID, Narration Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.AdvanceVouchersBodyBin b " & vbCrLf _
                  & " WHERE b.VoucherNo = '" & Grid.Columns("ID").Text & "'"
                  
   ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmRecoveryCustomerWise Then
      
      vStrSQL = " Select SerialNo, b.CustomerID ID, Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.RecoveryCustomerBin b " & vbCrLf _
                  & " WHERE b.RecoveryID = '" & Grid.Columns("ID").Text & "'"
                  
   ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmRecoveryInvoiceWise Then
      
      vStrSQL = " Select SerialNo, b.BillID ID, 'Bill Date ' + CONVERT (varchar, billdate,103 ) Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.RecoveryInvoiceBin b " & vbCrLf _
                  & " WHERE b.RecoveryID = '" & Grid.Columns("ID").Text & "'"
                  
   ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmPaymentInvoice Then
      
      vStrSQL = " Select SerialNo, b.PurID ID, 'Purchase Date ' + CONVERT (varchar, Purchasedate,103 ) Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.PaymentInvoiceBin b " & vbCrLf _
                  & " WHERE b.PaymentID = '" & Grid.Columns("ID").Text & "'"
                  
   ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmPaymentVender Then
      
      vStrSQL = " Select SerialNo, b.VenderID ID, Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.PaymentvenderBin b " & vbCrLf _
                  & " WHERE b.PaymentID = '" & Grid.Columns("ID").Text & "'"
                  
   ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmStockWastageInvoice Then
      
      vStrSQL = " Select SerialNo, b.WastageID ID, 'Qty:- ' + cast (floor(isnull(qtypack * multiplier,0) + qtyloose) as varchar) Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.StockWastageBodyBin b " & vbCrLf _
                  & " WHERE b.WastageID = '" & Grid.Columns("ID").Text & "' And WastageDate = '" & Grid.Columns("Date").Text & "'"
   ElseIf Trim(Grid.Columns("FormNo").Value) = eFrmStockAdjustment Then
      
      vStrSQL = " Select SerialNo, b.AdjustmentID ID, 'Qty:- ' + cast (floor(isnull(sqtypack * multiplier,0) + sqtyloose) as varchar) Description, amount " & vbCrLf _
                  & " from " & vBinDataBase & ".dbo.StockAdjustmentBodyBin b " & vbCrLf _
                  & " WHERE b.AdjustmentID = '" & Grid.Columns("ID").Text & "'"
                  
   End If
   
   If FrameDetail.Visible = False Then FrameDetail.Visible = True
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Set GridDetail.DataSource = Rs
   GridDetail.Columns("SerialNo").DataField = "SerialNo"
   GridDetail.Columns("ID").DataField = "ID"
   GridDetail.Columns("Description").DataField = "Description"
   GridDetail.Columns("Amount").DataField = "Amount"
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ReStoreData()
On Error GoTo ErrorHandler
   With Grid
      .Redraw = False
      .MoveFirst
         For vCounter = 1 To .Rows
            If Trim(.Columns("Selection").Value) = True And Trim(.Columns("ID").Value) <> "" Then
               CN.BeginTrans
               
                  ''''''''''''' Sale Invoice Recycle ''''''''''''''''
                  If Trim(.Columns("FormNo").Value) = eFrmSaleInvoicePOS Or Trim(.Columns("FormNo").Value) = eFrmSaleInvoiceDIS Then
                     
                     CN.Execute ("SET IDENTITY_INSERT SaleHeader ON")
                     vStrSQL = "Insert Into SaleHeader (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.SaleHeaderBin " & vbCrLf _
                        & "Where SID = " & .Columns("SID").Text
                     CN.Execute vStrSQL
                     vStrSQL = "Insert Into SaleBody (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.SaleBodyBin " & vbCrLf _
                        & "Where SID = " & .Columns("SID").Text
                     CN.Execute vStrSQL
                     CN.Execute ("SET IDENTITY_INSERT SaleHeader OFF")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.SaleBodyBin where SID = " & .Columns("SID").Text)
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.SaleHeaderBin where SID = " & .Columns("SID").Text)
                  
                  ''''''''''''' Sale Return Invoice Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmSaleReturnInvoicePOS Or Trim(.Columns("FormNo").Value) = eFrmSaleReturnInvoiceDIS Then
                     
                     CN.Execute ("SET IDENTITY_INSERT SaleReturnHeader ON")
                     vStrSQL = "Insert Into SaleReturnHeader (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.SaleReturnHeaderBin " & vbCrLf _
                        & "Where SID = " & .Columns("SID").Text
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into SaleReturnBody (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.SaleReturnBodyBin " & vbCrLf _
                                    & "Where SID = " & .Columns("SID").Text & " and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     CN.Execute ("SET IDENTITY_INSERT SaleReturnHeader OFF")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.SaleReturnBodyBin where SID = " & .Columns("SID").Text)
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.SaleReturnHeaderBin where SID = " & .Columns("SID").Text)
                  
                  ''''''''''''' Replacement Invoice Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmReplacementInvoice Then
                     
                     vSSID = CN.Execute("Select isnull(SSID,0) SSID from " & vBinDataBase & ".dbo.ReplacementHeaderBin where sid = " & .Columns("SID").Text).Fields(0).Value
                     vRSID = CN.Execute("Select isnull(RSID,0) RSID from " & vBinDataBase & ".dbo.ReplacementHeaderBin where sid = " & .Columns("SID").Text).Fields(0).Value
                     CN.Execute ("SET IDENTITY_INSERT ReplacementHeader ON")
                     vStrSQL = "Insert Into ReplacementHeader (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.ReplacementHeaderBin " & vbCrLf _
                        & "Where SID = " & .Columns("SID").Text
                     CN.Execute vStrSQL
                     CN.Execute ("SET IDENTITY_INSERT ReplacementHeader OFF")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.ReplacementHeaderBin where SID = " & .Columns("SID").Text)
                     
                     ''''''''''''' Replacement Sale Invoice Recycle ''''''''''''''''
                     CN.Execute ("SET IDENTITY_INSERT SaleHeader ON")
                     vStrSQL = "Insert Into SaleHeader (" & TableHeaderFields(eFrmSaleInvoicePOS) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(eFrmSaleInvoicePOS) & " from " & vBinDataBase & ".dbo.SaleHeaderBin " & vbCrLf _
                        & "Where SID = " & Val(vSSID)
                     CN.Execute vStrSQL
                     vStrSQL = "Insert Into SaleBody (" & TableBodyFields(eFrmSaleInvoicePOS) & ")" & vbCrLf _
                        & "Select " & TableBodyFields(eFrmSaleInvoicePOS) & " from  " & vBinDataBase & ".dbo.SaleBodyBin " & vbCrLf _
                        & "Where SID = " & Val(vSSID)
                     CN.Execute vStrSQL
                     CN.Execute ("SET IDENTITY_INSERT SaleHeader OFF")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.SaleBodyBin where SID = " & Val(vSSID))
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.SaleHeaderBin where SID = " & Val(vSSID))
                     
                     ''''''''''''' Replacement Sale Return Invoice Recycle ''''''''''''''''
                     CN.Execute ("SET IDENTITY_INSERT SaleReturnHeader ON")
                     vStrSQL = "Insert Into SaleReturnHeader (" & TableHeaderFields(eFrmSaleReturnInvoicePOS) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(eFrmSaleReturnInvoicePOS) & " from " & vBinDataBase & ".dbo.SaleReturnHeaderBin " & vbCrLf _
                        & "Where SID = " & Val(vRSID)
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Select SerialNo from " & vBinDataBase & ".dbo.SaleReturnBodyBin where sid = " & Val(vRSID) & " and SerialNo = '" & GridDetail.Columns("SerialNo").Text & "'"
                           If CN.Execute(vStrSQL).EOF = False Then
                           
                              vStrSQL = "Insert Into SaleReturnBody (" & TableBodyFields(eFrmSaleReturnInvoicePOS) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(eFrmSaleReturnInvoicePOS) & " from  " & vBinDataBase & ".dbo.SaleReturnBodyBin " & vbCrLf _
                                    & "Where SID = " & Val(vRSID) & " and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                              CN.Execute vStrSQL
                           End If
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     CN.Execute ("SET IDENTITY_INSERT SaleReturnHeader OFF")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.SaleReturnBodyBin where SID = " & Val(vRSID))
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.SaleReturnHeaderBin where SID = " & Val(vRSID))
                  
                  ''''''''''''' Sale Order Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmSaleOrderPOS Or Trim(.Columns("FormNo").Value) = eFrmSaleOrderDIS Then
                     vStrSQL = "Insert Into SaleOrderHeader (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.SaleOrderHeaderBin " & vbCrLf _
                        & " WHERE Orderid = '" & Grid.Columns("ID").Text & "' And OrderDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into SaleOrderBody (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.SaleOrderBodyBin " & vbCrLf _
                                    & " WHERE Orderid = '" & Grid.Columns("ID").Text & "' And OrderDate = '" & Grid.Columns("Date").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.SaleOrderBodyBin WHERE Orderid = '" & .Columns("ID").Text & "' And OrderDate = '" & .Columns("Date").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.SaleOrderHeaderBin WHERE Orderid = '" & .Columns("ID").Text & "' And OrderDate = '" & .Columns("Date").Text & "'")
                     
                  ''''''''''''' Purchase Invoice Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmPurchaseInvoice Then
                     
                     vStrSQL = "Insert Into PurchaseHeader (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.PurchaseHeaderBin " & vbCrLf _
                        & " WHERE purid = '" & Grid.Columns("ID").Text & "' And PurchaseDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into PurchaseBody (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.PurchaseBodyBin " & vbCrLf _
                                    & " WHERE purid = '" & Grid.Columns("ID").Text & "' And PurchaseDate = '" & Grid.Columns("Date").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.PurchaseBodyBin WHERE purid = '" & .Columns("ID").Text & "' And PurchaseDate = '" & .Columns("Date").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.PurchaseHeaderBin WHERE purid = '" & .Columns("ID").Text & "' And PurchaseDate = '" & .Columns("Date").Text & "'")
                                          
                  ''''''''''''' Purchase Order Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmPurchaseOrder Then
                     vStrSQL = "Insert Into PurchaseOrderHeader (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.PurchaseOrderHeaderBin " & vbCrLf _
                        & " WHERE Orderid = '" & Grid.Columns("ID").Text & "' And OrderDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into PurchaseOrderBody (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.PurchaseOrderBodyBin " & vbCrLf _
                                    & " WHERE Orderid = '" & Grid.Columns("ID").Text & "' And OrderDate = '" & Grid.Columns("Date").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.PurchaseOrderBodyBin WHERE Orderid = '" & .Columns("ID").Text & "' And OrderDate = '" & .Columns("Date").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.PurchaseOrderHeaderBin WHERE Orderid = '" & .Columns("ID").Text & "' And OrderDate = '" & .Columns("Date").Text & "'")
                     
                     ''''''''''''' Purchase Return Invoice Recycle ''''''''''''''''
                     ElseIf Trim(.Columns("FormNo").Value) = eFrmPurchaseReturnInvoice Then
                                          
                     vStrSQL = "Insert Into PurchaseReturnHeader (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.PurchaseReturnHeaderBin " & vbCrLf _
                        & " WHERE ReturnID = '" & Grid.Columns("ID").Text & "' And ReturnDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into PurchaseReturnBody (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.PurchaseReturnBodyBin " & vbCrLf _
                                    & " WHERE ReturnID = '" & Grid.Columns("ID").Text & "' And ReturnDate = '" & Grid.Columns("Date").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.PurchaseReturnBodyBin WHERE ReturnID = '" & .Columns("ID").Text & "' And ReturnDate = '" & .Columns("Date").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.PurchaseReturnHeaderBin WHERE ReturnID = '" & .Columns("ID").Text & "' And ReturnDate = '" & .Columns("Date").Text & "'")
                  
                  ''''''''''''' Credit Vouchers Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmCreditVoucher Then
                                          
                     vStrSQL = "Insert Into CreditVouchers (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.CreditVouchersBin " & vbCrLf _
                        & " WHERE VoucherNo = '" & Grid.Columns("ID").Text & "' And VoucherDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into CreditVouchersBody (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.CreditVouchersBodyBin " & vbCrLf _
                                    & " WHERE VoucherNo = '" & Grid.Columns("ID").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.CreditVouchersBodyBin WHERE VoucherNo = '" & .Columns("ID").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.CreditVouchersBin WHERE VoucherNo = '" & .Columns("ID").Text & "'")
                     
                  ''''''''''''' Debit Vouchers Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmDebitVoucher Then
                     
                     vStrSQL = "Insert Into DebitVouchers (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.DebitVouchersBin " & vbCrLf _
                        & " WHERE VoucherNo = '" & Grid.Columns("ID").Text & "' And VoucherDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into DebitVouchersBody (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.DebitVouchersBodyBin " & vbCrLf _
                                    & " WHERE VoucherNo = '" & Grid.Columns("ID").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.DebitVouchersBodyBin WHERE VoucherNo = '" & .Columns("ID").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.DebitVouchersBin WHERE VoucherNo = '" & .Columns("ID").Text & "'")
                  
                  ''''''''''''' Journal Vouchers Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmJournalVoucher Then
                                       
                     vStrSQL = "Insert Into JournalVouchers (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.JournalVouchersBin " & vbCrLf _
                        & " WHERE VoucherNo = '" & Grid.Columns("ID").Text & "' And VoucherDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into JournalVouchersBody (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.JournalVouchersBodyBin " & vbCrLf _
                                    & " WHERE VoucherNo = '" & Grid.Columns("ID").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.JournalVouchersBodyBin WHERE VoucherNo = '" & .Columns("ID").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.JournalVouchersBin WHERE VoucherNo = '" & .Columns("ID").Text & "'")
                  
                  ''''''''''''' Advance Vouchers Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmAdvanceVoucher Then
                                          
                     vStrSQL = "Insert Into AdvanceVouchers (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.AdvanceVouchersBin " & vbCrLf _
                        & " WHERE VoucherNo = '" & Grid.Columns("ID").Text & "' And VoucherDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into AdvanceVouchersBody (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.AdvanceVouchersBodyBin " & vbCrLf _
                                    & " WHERE VoucherNo = '" & Grid.Columns("ID").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.AdvanceVouchersBodyBin WHERE VoucherNo = '" & .Columns("ID").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.AdvanceVouchersBin WHERE VoucherNo = '" & .Columns("ID").Text & "'")
                     
                  ''''''''''''' RecoveryCustomerWise Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmRecoveryCustomerWise Then
                     
                     vStrSQL = "Insert Into RecoveryHeader (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.RecoveryHeaderBin " & vbCrLf _
                        & " WHERE RecoveryID = '" & Grid.Columns("ID").Text & "' And RecoveryDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into RecoveryCustomer (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.RecoveryCustomerBin " & vbCrLf _
                                    & " WHERE RecoveryID = '" & Grid.Columns("ID").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.RecoveryCustomerBin WHERE RecoveryID = '" & .Columns("ID").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.RecoveryHeaderBin WHERE RecoveryID = '" & .Columns("ID").Text & "'")
                     
               ''''''''''''' RecoveryInvoiceWise Recycle ''''''''''''''''
               ElseIf Trim(.Columns("FormNo").Value) = eFrmRecoveryInvoiceWise Then
                     
                     vStrSQL = "Insert Into RecoveryHeader (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.RecoveryHeaderBin " & vbCrLf _
                        & " WHERE RecoveryID = '" & Grid.Columns("ID").Text & "' And RecoveryDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into RecoveryInvoice (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.RecoveryInvoiceBin " & vbCrLf _
                                    & " WHERE RecoveryID = '" & Grid.Columns("ID").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.RecoveryInvoiceBin WHERE RecoveryID = '" & .Columns("ID").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.RecoveryHeaderBin WHERE RecoveryID = '" & .Columns("ID").Text & "'")
               
                ''''''''''''' Payment Invoice Recycle ''''''''''''''''
                ElseIf Trim(.Columns("FormNo").Value) = eFrmPaymentInvoice Then
                     
                     vStrSQL = "Insert Into PaymentHeader (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.PaymentHeaderBin " & vbCrLf _
                        & " WHERE PaymentID = '" & Grid.Columns("ID").Text & "' And PaymentDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into PaymentInvoice (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.PaymentInvoiceBin " & vbCrLf _
                                    & " WHERE PaymentID = '" & Grid.Columns("ID").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.PaymentInvoiceBin WHERE PaymentID = '" & .Columns("ID").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.PaymentHeaderBin WHERE PaymentID = '" & .Columns("ID").Text & "'")
               
               ''''''''''''' Payment Vender Recycle ''''''''''''''''
               ElseIf Trim(.Columns("FormNo").Value) = eFrmPaymentVender Then
                     
                     vStrSQL = "Insert Into PaymentHeader (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.PaymentHeaderBin " & vbCrLf _
                        & " WHERE PaymentID = '" & Grid.Columns("ID").Text & "' And PaymentDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into PaymentVender (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.PaymentVenderBin " & vbCrLf _
                                    & " WHERE PaymentID = '" & Grid.Columns("ID").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.PaymentVenderBin WHERE PaymentID = '" & .Columns("ID").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.PaymentHeaderBin WHERE PaymentID = '" & .Columns("ID").Text & "'")
                  
                  ''''''''''''' Stock Wastage Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmStockWastageInvoice Then
                     
                     vStrSQL = "Insert Into StockWastageHeader (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.StockWastageHeaderBin " & vbCrLf _
                        & " WHERE WastageID = '" & Grid.Columns("ID").Text & "' And WastageDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into StockWastageBody (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.StockWastageBodyBin " & vbCrLf _
                                    & " WHERE WastageID = '" & Grid.Columns("ID").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.StockWastageBodyBin WHERE WastageID = '" & .Columns("ID").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.StockWastageHeaderBin WHERE WastageID = '" & .Columns("ID").Text & "'")
                  
                  ''''''''''''' Stock Adjustment Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmStockAdjustment Then
                     
                     vStrSQL = "Insert Into StockAdjustmentHeader (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.StockAdjustmentHeaderBin " & vbCrLf _
                        & " WHERE AdjustmentID = '" & Grid.Columns("ID").Text & "' And AdjustmentDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     
                     If GridDetail.Rows > 0 Then
                        GridDetail.Redraw = False
                        GridDetail.MoveFirst
                        For i = 1 To GridDetail.Rows
                           vStrSQL = "Insert Into StockAdjustmentBody (" & TableBodyFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                                    & "Select " & TableBodyFields(.Columns("FormNo").Value) & " from  " & vBinDataBase & ".dbo.StockAdjustmentBodyBin " & vbCrLf _
                                    & " WHERE AdjustmentID = '" & Grid.Columns("ID").Text & "' and SerialNo = " & GridDetail.Columns("SerialNo").Text
                                    
                           CN.Execute vStrSQL
                           If i < GridDetail.Rows Then GridDetail.MoveNext
                        Next i
                        GridDetail.Redraw = True
                     End If
                     
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.StockAdjustmentBodyBin WHERE AdjustmentID = '" & .Columns("ID").Text & "'")
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.StockAdjustmentHeaderBin WHERE AdjustmentID = '" & .Columns("ID").Text & "'")
                    
                  ''''''''''''' AdminClosing Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmAdminClosing Then
                     
                     vStrSQL = "Insert Into AdminClosing (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.AdminClosingBin " & vbCrLf _
                        & " WHERE ID = '" & Grid.Columns("ID").Text & "' And EntryDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.AdminClosingBin WHERE ID = '" & .Columns("ID").Text & "' And EntryDate = '" & Grid.Columns("Date").Text & "'")
                     
                     
                  ''''''''''''' Salaries Recycle ''''''''''''''''
                  ElseIf Trim(.Columns("FormNo").Value) = eFrmSalaries Then
                     
                     vStrSQL = "Insert Into Salaries (" & TableHeaderFields(.Columns("FormNo").Value) & ")" & vbCrLf _
                        & "Select " & TableHeaderFields(.Columns("FormNo").Value) & " from " & vBinDataBase & ".dbo.SalariesBin " & vbCrLf _
                        & " WHERE SalaryID = '" & Grid.Columns("ID").Text & "' And EntryDate = '" & Grid.Columns("Date").Text & "'"
                     CN.Execute vStrSQL
                     CN.Execute ("Delete " & vBinDataBase & ".dbo.SalariesBin WHERE SalaryID = '" & .Columns("ID").Text & "' And EntryDate = '" & Grid.Columns("Date").Text & "'")
                      
                  End If
                  ''''''''''''''''''''''''''''''''''''''''''''''''''''
               
               vStrSQL = "insert into " & vBinDataBase & ".dbo.ActivityLogBin(ActivityDate,ActionNo,userno,FormNo,TransactionID,TransactionDate,TransactionInfo) values(getdate()," & eReStoreRecord & "," & 1 & ",'" & .Columns("FormNo").Value & "'," & .Columns("ID").Value & ", '" & .Columns("Date").Value & "','" & IIf(GridDetail.Columns("ID").Text = "", "Data", GridDetail.Rows & " Row/s") & " ReCycled')"
               CN.Execute vStrSQL
               CN.CommitTrans
            End If
         .MoveNext
         Next vCounter
         .Redraw = True
      End With
   LoadGrid
Exit Sub
ErrorHandler:
   Grid.Redraw = True
   GridDetail.Redraw = True
   If eFrmSaleInvoicePOS = Grid.Columns("FormNo").Text Or eFrmSaleInvoiceDIS = Grid.Columns("FormNo").Text Then CN.Execute ("SET IDENTITY_INSERT SaleHeader OFF")
   If eFrmSaleReturnInvoicePOS = Grid.Columns("FormNo").Text Or eFrmSaleReturnInvoiceDIS = Grid.Columns("FormNo").Text Then CN.Execute ("SET IDENTITY_INSERT SaleReutrnHeader OFF")
   If eFrmReplacementInvoice = Grid.Columns("FormNo").Text Then
      CN.Execute ("SET IDENTITY_INSERT ReplacementHeader OFF")
      CN.Execute ("SET IDENTITY_INSERT SaleHeader OFF")
      CN.Execute ("SET IDENTITY_INSERT SaleReutrnHeader OFF")
   End If
   If CN.Errors.Count > 0 Then CN.RollbackTrans
   Call ShowErrorMessage
End Sub



