VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmRecycleBin 
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
   Begin VB.ComboBox CmbAction 
      Height          =   315
      ItemData        =   "FrmRecycleBin.frx":0000
      Left            =   10958
      List            =   "FrmRecycleBin.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2160
      Width           =   2250
   End
   Begin VB.ComboBox CmbUser 
      Height          =   315
      Left            =   2048
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2160
      Width           =   2790
   End
   Begin VB.ComboBox CmbFilter 
      Height          =   315
      Left            =   4883
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2160
      Width           =   3150
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7051
      TabIndex        =   6
      Top             =   8730
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
      MICON           =   "FrmRecycleBin.frx":0004
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   4883
      TabIndex        =   5
      Top             =   8775
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Select"
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
      MICON           =   "FrmRecycleBin.frx":0020
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFromDate 
      Height          =   330
      Left            =   8063
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
      Left            =   9503
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
      CausesValidation=   0   'False
      Height          =   3270
      Left            =   1868
      TabIndex        =   12
      Top             =   2790
      Width           =   11625
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   6
      stylesets.count =   4
      stylesets(0).Name=   "Red"
      stylesets(0).ForeColor=   665589
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "FrmRecycleBin.frx":003C
      stylesets(1).Name=   "Select"
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
      stylesets(1).Picture=   "FrmRecycleBin.frx":0058
      stylesets(2).Name=   "Orange"
      stylesets(2).ForeColor=   33023
      stylesets(2).HasFont=   -1  'True
      BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(2).Picture=   "FrmRecycleBin.frx":0074
      stylesets(3).Name=   "Green"
      stylesets(3).ForeColor=   2135858
      stylesets(3).HasFont=   -1  'True
      BeginProperty stylesets(3).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(3).Picture=   "FrmRecycleBin.frx":0090
      AllowUpdate     =   0   'False
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
      RowNavigation   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   503
      ExtraHeight     =   767
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   6
      Columns(0).Width=   1005
      Columns(0).Caption=   "ID"
      Columns(0).Name =   "ID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2566
      Columns(1).Caption=   "Date"
      Columns(1).Name =   "Date"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2593
      Columns(2).Caption=   "UserName"
      Columns(2).Name =   "UserName"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   4180
      Columns(3).Caption=   "Voucher Info"
      Columns(3).Name =   "VoucherInfo"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   7699
      Columns(4).Caption=   "Detail Info"
      Columns(4).Name =   "Description"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1984
      Columns(5).Caption=   "Action"
      Columns(5).Name =   "Action"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20505
      _ExtentY        =   5768
      _StockProps     =   79
      BackColor       =   15724527
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   10958
      TabIndex        =   11
      Top             =   1935
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
      Left            =   2048
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
      Left            =   4883
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
      Left            =   8153
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
Attribute VB_Name = "FrmRecycleBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim RsBody As New ADODB.Recordset
Dim vStrSQL As String
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

Private Sub CmbAction_Click()
   On Error GoTo ErrorHandler
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
   If CmbFilter.ListIndex = 0 Then
      vFilter = ""
   Else
'      vFilter = " and FormType like '%" & CmbFilter.Text & "%'"
   End If
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbUser_Click()
   On Error GoTo ErrorHandler
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
   SetWindowText Me.hwnd, "Activity Log"
   CmbFilter.Clear
   
   CmbFilter.AddItem "Sale Invoice"
          
  
   CmbUser.Clear
   CmbUser.AddItem "-- All Users --"
   With CN.Execute("Select * FROM Users")
      Do Until .EOF
          CmbUser.AddItem !UserName
          CmbUser.ItemData(CmbUser.NewIndex) = !UserNo
          .MoveNext
      Loop
   End With
   
   CmbAction.Clear
   CmbAction.AddItem "-- All Actions --"
'   CmbAction.AddItem "New"
'   CmbAction.AddItem "Edit"
   CmbAction.AddItem "Delete"
   CmbAction.AddItem "Clear"
   CmbAction.ListIndex = 0
   
   DtpFromDate.DateValue = Date - 10
   DtpToDate.DateValue = Date
   CmbFilter.ListIndex = 0
   CmbUser.ListIndex = 0
   vOrder = "order by billdate Desc"
   LoadGrid
   Me.ParaOutUserNo = ""
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   vStrSQL = " Select H.billid, h.billdate, isnull(h.remarks,'') remarks, isnull(h.description,'') Description, UserName from Bin_SaleHeader H " & _
             " Inner Join Users U on U.UserNO= H.UserNo  " & " WHERE h.billdate between '" & DtpFromDate.DateValue & "' and '" & DateAdd("d", 1, DtpToDate.DateValue) & "'" & vAction & vFilter & vUser & vOrder & vDirection
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
   Grid.Redraw = False
   Grid.MoveFirst
   Grid.RemoveAll
   Grid.AllowAddNew = True
   While Not Rs.EOF
      Grid.Columns("ID").Text = Rs!BillID
      Grid.Columns("Date").Text = Rs!billdate
      Grid.Columns("VoucherInfo").Text = Rs!Description
      Grid.Columns("Action").Text = Rs!Remarks
      Grid.Columns("UserName").Text = Rs!UserName
      vStrSQL = " Select b.productId, ProductName, isnull(PackingName,'') PackingName, QtyPack,  qty, amount " & _
                " from Bin_SaleBody b  Left Outer join products p on p.productid = b.productid " & _
                " left outer join Packings pa on pa.packingid = b.packingid " & _
                " WHERE b.billid = '" & Rs!BillID & "' and b.billDate = '" & Rs!billdate & "'"
      If RsBody.State = adStateOpen Then RsBody.Close
      RsBody.Open vStrSQL, CN, adOpenStatic, adLockReadOnly
      While Not RsBody.EOF
         Grid.Columns("Description").Text = Grid.Columns("Description").Text & " " & RsBody!ProductName & " " & RsBody!Amount
         RsBody.MoveNext
      Wend
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
   vOrder = " order by " & Grid.Columns(ColIndex).DataField
   If vCol = ColIndex Then
      vDirection = IIf(vDirection = " Asc", " Desc", " Asc")
   Else
      vDirection = " Asc"
   End If
   vCol = ColIndex
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   On Error GoTo ErrorHandler
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

