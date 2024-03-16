VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form FrmPostSelectedSales 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmPostSelectedSales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   1  'CenterOwner
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   4095
      Left            =   1890
      TabIndex        =   0
      Top             =   3375
      Width           =   11715
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   14
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
      stylesets(0).Picture=   "FrmPostSelectedSales.frx":0ECA
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
      stylesets(1).Picture=   "FrmPostSelectedSales.frx":0EE6
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "SelectedCol"
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
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
      Columns.Count   =   14
      Columns(0).Width=   1138
      Columns(0).Caption=   "ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   1984
      Columns(1).Caption=   "Date"
      Columns(1).Name =   "Date"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1720
      Columns(2).Caption=   "BillTime"
      Columns(2).Name =   "BillTime"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "StoreID"
      Columns(3).Name =   "StoreID"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2566
      Columns(4).Caption=   "StoreName"
      Columns(4).Name =   "StoreName"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1588
      Columns(5).Caption=   "TotalItems"
      Columns(5).Name =   "TotalItems"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "CustomerID"
      Columns(6).Name =   "CustomerID"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3810
      Columns(7).Caption=   "CustomerName"
      Columns(7).Name =   "CustomerName"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1085
      Columns(8).Caption=   "Disc"
      Columns(8).Name =   "Disc"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1535
      Columns(9).Caption=   "BillType"
      Columns(9).Name =   "BillType"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1640
      Columns(10).Caption=   "Amount"
      Columns(10).Name=   "Amount"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   1270
      Columns(11).Caption=   "CO"
      Columns(11).Name=   "CO"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   1164
      Columns(12).Caption=   "Post"
      Columns(12).Name=   "Lock"
      Columns(12).CaptionAlignment=   2
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(12).Style=   2
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "SID"
      Columns(13).Name=   "SID"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   20664
      _ExtentY        =   7223
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
   Begin JeweledBut.JeweledButton BtnDelete 
      Height          =   420
      Left            =   6293
      TabIndex        =   2
      Top             =   8565
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
      MICON           =   "FrmPostSelectedSales.frx":0F02
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7793
      TabIndex        =   3
      Top             =   8565
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
      MICON           =   "FrmPostSelectedSales.frx":0F1E
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFromDate 
      Height          =   330
      Left            =   2625
      TabIndex        =   4
      Top             =   3000
      Width           =   1200
      _Version        =   65543
      _ExtentX        =   2117
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
      Left            =   3825
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   1200
      _Version        =   65543
      _ExtentX        =   2117
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
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
      Left            =   1845
      TabIndex        =   6
      Top             =   3090
      Width           =   735
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Post Selected Sale Invoice"
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
      TabIndex        =   1
      Top             =   270
      Width           =   3495
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "FrmPostSelectedSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rs As New ADODB.Recordset
Public vSuppressUpdateEvent As Boolean
Public ParaOutID As String
Dim vMargin As String
Dim vCounter As Integer, FunGetMaxID As Long
Dim vStrSQL, sSql, vConstraintName, vForignKey As String



Private Sub BtnMemberDisc_Click()
'   FrmMembersDiscount.Show vbModal
End Sub

Private Sub BtnClose_Click()
  Unload Me
End Sub



Private Sub BtnDelete_Click()
   On Error GoTo ErrorHandler
   CN.BeginTrans
   Grid.Redraw = False
   Grid.MoveFirst
   For vCounter = 1 To Grid.Rows
   
              
      ''''''''''''' Delete Sale Header ''''''''''''''''''''''''
      vStrSQL = "update  SaleHeader set isPosted = " & Abs(Grid.Columns("Lock").Value) & " Where SID = " & Grid.Columns("SID").Value & " and billID = " & Grid.Columns("ID").Value & " and BillDate ='" & Grid.Columns("Date").Text & "'"
      CN.Execute (vStrSQL)
   
      Grid.MoveNext
   Next vCounter
   Grid.RemoveAll
   Grid.Redraw = True
   CN.CommitTrans
   Call LoadGrid
   

   Exit Sub
ErrorHandler:
  Grid.Redraw = True
  Call ShowErrorMessage
End Sub

Private Function GetGroupID(cmb As ComboBox) As String
    On Error GoTo ErrorHandler
    If cmb.ListIndex < 0 Then Exit Function
    GetGroupID = Chr(Left(cmb.ItemData(cmb.ListIndex), 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 3, 2)) & Chr(Mid(cmb.ItemData(cmb.ListIndex), 5, 2))
    Exit Function
ErrorHandler:
    Call ShowErrorMessage
End Function

Private Sub CmbSortBy_Click()
   On Error GoTo ErrorHandler
'   PopulateGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub DtpFromDate_Change()
   If DtpFromDate.IsDateValid = False Then Exit Sub
   If DtpToDate.Visible = False Then DtpToDate.DateValue = DtpFromDate.DateValue
   Call LoadGrid
End Sub

Private Sub DtpToDate_Change()
   If DtpToDate.IsDateValid = False Then Exit Sub
   Call LoadGrid
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hwnd, "Change Price"
   Call LoadGrid
'   Call PopulateGrid
   BtnDelete.Visible = Not ObjRegistry.ReadOnlyStatus
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name <> Grid.Name Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      End If
   ElseIf Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnDelete.Enabled Then BtnDelete_Click
            KeyCode = 0
         Case vbKeyW
'            If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   ElseIf Shift = 0 And KeyCode <> 0 Then
      If UCase(Me.ActiveControl.Name) Like "TXT*" Then If BtnDelete.Enabled = False Then BtnDelete.Enabled = True
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
'   If Button = 1 Then
'      Call ReleaseCapture
'      lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
'   End If
End Sub


Private Sub ImgExit_Click()
   Unload Me
End Sub


Private Sub PopulateGrid()
   On Error GoTo ErrorHandler
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
   'CmbClass.ListIndex = 0
   Me.MousePointer = vbHourglass
   
      sSql = " Select H.syllabusID, H.SyllabusName, b.code, b.productid, Productname, QtyLoose, Price, Amount, isShow  from syllabusheader H " & vbCrLf _
      + "inner join Schools S on S.SchoolID = H.SchoolID Inner join Classes C on C.ClassID = H.ClassID  " & vbCrLf _
       + " inner join SyllabusBody b on b.syllabusID = h.syllabusID inner join Products P on P.ProductID = b.ProductID Where S.SchoolID = " & 1
   
   With CN.Execute(sSql)
      Do Until .EOF
         Grid.AddNew
         Grid.Columns("ID").Text = !ProductID
         Grid.Columns("Name").Text = !ProductName
         Grid.Columns("Price").Value = !Price
         Grid.Columns("Qty").Value = !QtyLoose
         Grid.Columns("Amount").Value = !Amount
         Grid.Columns("Lock").Value = !isShow
         Grid.Update
         .MoveNext
      Loop
   End With
   If Rs.State = adStateOpen Then
     Rs.CancelBatch
     Rs.Close
   End If
   Rs.Open "Select * FROM SyllabusBody where 1=1 and SyllabusID = " & 1, CN, adOpenStatic, adLockBatchOptimistic
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
Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   Grid.Redraw = False
   Grid.CancelUpdate
   Grid.RemoveAll
   vSuppressUpdateEvent = True
   'CmbClass.ListIndex = 0
   Me.MousePointer = vbHourglass
   vStrSQL = "SELECT h.Sid, h.BillID as SaleID, h.StoreID, h.isPosted, c.AccountNo, OrderID, h.BillDate as SaleDate, TableName, Substring(CONVERT(varchar(20),isnull(BillTime,0)),13,7) as BillTime, case when credit = 1 then 'Credit' when cash = 1 then 'Cash' when BankCard = 1 then 'Bank Card' end as BillType, InvType" & vbCrLf _
         + " , Case when CustomerID = '621' then isnull(CustomerName,c.AccountName) Else AccountName + isnull(' (' + City + ')','') + isnull(' (' + address + ')','') End as CustomerName, TotalAmount-isnull(billdisc,0)+isnull(ServiceCharges,0)+isnull(STax,0)+isnull(othercharges,0) as TotalAmount, isnull(billdisc,0) + disc as Disc, isnull(ServiceCharges,0) as SC, TotalItems, UserName, isPosted, isReplace, StoreName, Tag, isnull(ManualBillNo,'')ManualBillNo" & vbCrLf _
         + " FROM SaleHeader h INNER JOIN" & vbCrLf _
         + " (SELECT SID, BillDate, sum(isnull(multiplier,0)* isnull(QtyPack,0) + Qty + isnull(Bonus,0)) as TotalItems, sum(DiscVal) as disc, sum(amount) Amount FROM SaleBody GROUP BY SID, BillDate) b" & vbCrLf _
         + " ON h.SID = b.SID and h.BillDate = b.BillDate" & vbCrLf _
         + " left outer JOIN chartofaccounts c ON h.CustomerID = c.AccountNo " & vbCrLf _
         + " left outer JOIN Parties p ON p.PartyID = c.AccountNo " & vbCrLf _
         + " left outer JOIN Tables tb ON tb.TableID = h.TableID " & vbCrLf _
         + " INNER JOIN users u ON h.userno = u.userno " & vbCrLf _
         + " INNER JOIN Stores s ON s.StoreID = h.StoreID " & vbCrLf _
         + " WHERE h.BillDate = '" & DtpFromDate.DateValue & "'"
'         + " WHERE h.BillDate Between '" & DtpFromDate.DateValue & "' and '" & DtpToDate.DateValue & "'" ' & vPartyName & vTag & vTableName & vType & vTtlAmount & vManualBillNo & vOrder & vDirection
   With CN.Execute(vStrSQL)
      Do Until .EOF
         Grid.AddNew
         Grid.Columns("SID").Text = !SID
         Grid.Columns("ID").Text = !SaleID
         Grid.Columns("Date").Text = !SaleDate
         Grid.Columns("BillTime").Text = !BillTime
         Grid.Columns("CustomerID").Text = !AccountNo
         Grid.Columns("CustomerName").Text = !CustomerName
         Grid.Columns("TotalItems").Value = !TotalItems
         Grid.Columns("Disc").Value = !Disc
         Grid.Columns("Amount").Value = !TotalAmount
         Grid.Columns("CO").Text = !UserName
         Grid.Columns("Lock").Text = !isPosted
         Grid.Columns("BillType").Text = !BillType
         Grid.Columns("StoreID").Text = !StoreID
         Grid.Columns("StoreName").Text = !StoreName
         Grid.Update
         .MoveNext
      Loop
   End With
   vSuppressUpdateEvent = False
   Grid.Redraw = True
   Grid.MoveFirst
   Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
    Grid.Redraw = True
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub
