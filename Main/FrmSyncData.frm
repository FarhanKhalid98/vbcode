VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form FrmSyncData 
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LVDef 
      Height          =   3840
      Left            =   6750
      TabIndex        =   1
      Tag             =   "C"
      ToolTipText     =   "Product Entry"
      Top             =   1350
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   6773
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
   Begin JeweledBut.JeweledButton BtnRefresh 
      Height          =   420
      Left            =   2340
      TabIndex        =   2
      Top             =   6930
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Refresh"
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
      MICON           =   "FrmSyncData.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin MSComctlLib.ListView LVTrans 
      Height          =   4290
      Left            =   3375
      TabIndex        =   3
      Tag             =   "C"
      ToolTipText     =   "Product Entry"
      Top             =   1350
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   7567
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridTable 
      CausesValidation=   0   'False
      Height          =   3840
      Left            =   10125
      TabIndex        =   4
      Top             =   1350
      Visible         =   0   'False
      Width           =   2835
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   2
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
      stylesets(0).Picture=   "FrmSyncData.frx":001C
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
      stylesets(1).Picture=   "FrmSyncData.frx":0038
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
      stylesets(2).Picture=   "FrmSyncData.frx":0054
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
      stylesets(3).Picture=   "FrmSyncData.frx":0070
      AllowUpdate     =   0   'False
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
      RowHeight       =   529
      ExtraHeight     =   423
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   2
      Columns(0).Width=   3863
      Columns(0).Caption=   "Column_Name"
      Columns(0).Name =   "Column_Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Data_Type"
      Columns(1).Name =   "Data_Type"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   5001
      _ExtentY        =   6773
      _StockProps     =   79
      BackColor       =   15724527
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   9465
      TabIndex        =   5
      Top             =   6870
      Width           =   1500
      _ExtentX        =   2646
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
      MICON           =   "FrmSyncData.frx":008C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnTransactionImport 
      Height          =   420
      Left            =   6120
      TabIndex        =   6
      Top             =   6885
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   741
      TX              =   "Transaction Import"
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
      MICON           =   "FrmSyncData.frx":00A8
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnDefinationExport 
      Height          =   420
      Left            =   7800
      TabIndex        =   7
      Top             =   6885
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   741
      TX              =   "Defination Export"
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
      MICON           =   "FrmSyncData.frx":00C4
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnStockTrnsfer 
      Height          =   420
      Left            =   4455
      TabIndex        =   8
      Top             =   6885
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   741
      TX              =   "Stock Transfer Export"
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
      MICON           =   "FrmSyncData.frx":00E0
      BC              =   14737632
      FC              =   0
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid GridStore 
      Height          =   3840
      Left            =   0
      TabIndex        =   9
      Top             =   1350
      Visible         =   0   'False
      Width           =   2835
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
      stylesets(0).Picture=   "FrmSyncData.frx":00FC
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
      RowHeight       =   714
      ActiveRowStyleSet=   "Select"
      Columns.Count   =   3
      Columns(0).Width=   3200
      Columns(0).Caption=   "StoreID"
      Columns(0).Name =   "StoreID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Config"
      Columns(1).Name =   "Config"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "AccountNo"
      Columns(2).Name =   "AccountNo"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   5001
      _ExtentY        =   6773
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
   Begin JeweledBut.JeweledButton BtnSyncAll 
      Height          =   420
      Left            =   7080
      TabIndex        =   10
      Top             =   8010
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Sync All"
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
      MICON           =   "FrmSyncData.frx":0118
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClearSyncDefination 
      Height          =   420
      Left            =   6750
      TabIndex        =   11
      Top             =   5190
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   741
      TX              =   "Clear Sync Defination"
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
      MICON           =   "FrmSyncData.frx":0134
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sync Data"
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
      TabIndex        =   0
      Top             =   270
      Width           =   1410
   End
End
Attribute VB_Name = "FrmSyncData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSQL, vColumnList1, vColumnList2, vColumnList3, vwhere, vJoin, vPKey1, vPKey2, vPKey3, vSerialNo, vHeaderTable, vConfig As String
Dim Item As ListItem
Dim i, vPKeyCount, vIdentityKey, vStoreID, vAccountNo, vTableCount As Integer
Dim FunGetMaxID As Integer
Dim vPrice, vMultiplier, vAmount, vTotalAmount, vSID As Double
Dim Rs As New ADODB.Recordset
Dim vLinkedServer As String


Private Sub BtnClearSyncDefination_Click()
On Error GoTo ErrorHandler

      Me.MousePointer = vbHourglass
      For vTableCount = 1 To LVDef.ListItems.Count
         If LVDef.ListItems(vTableCount).Checked = True Then
            vSQL = "Update " & LVDef.ListItems(vTableCount).Text & " Set IsSync = 0 where IsSync = 1 "
            CN.Execute vSQL
'            LVDef.ListItems(vTableCount).Checked = False
         End If
      Next vTableCount
         MsgBox "Clear Sync Defination Succeed", vbInformation, Me.Caption
         Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnDefinationExport_Click()
 On Error GoTo ErrorHandler
      Me.MousePointer = vbHourglass
      vSQL = "Select * from Stores where StoreID <> 1 and isLock = 0 and Config is not null"
      If Rs.State = adStateOpen Then Rs.Close
      Rs.Open vSQL, CN, adOpenStatic, adLockReadOnly
   
      While Not Rs.EOF
         vStoreID = Rs!StoreID
         vConfig = Rs!Config
         Call DefinationExport
         Rs.MoveNext
      Wend
      
      For vTableCount = 1 To LVDef.ListItems.Count
         If LVDef.ListItems(vTableCount).Checked = True Then
            vSQL = "Update " & LVDef.ListItems(vTableCount).Text & " Set IsSync = 1 where IsSync = 0 "
            CN.Execute vSQL
'            LVDef.ListItems(vTableCount).Checked = False
         End If
      Next vTableCount
      
         MsgBox "Sync Succeed", vbInformation, Me.Caption
         Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnRefresh_Click()
   Call Settings
End Sub

Private Sub BtnStockTrnsfer_Click()
On Error GoTo ErrorHandler
      Me.MousePointer = vbHourglass
      
         vSQL = "Select * from Stores where StoreID <> 1 and isLock = 0 and Config is not null"
         If Rs.State = adStateOpen Then Rs.Close
         Rs.Open vSQL, CN, adOpenStatic, adLockReadOnly
         
         Set GridStore.DataSource = Rs
         GridStore.Columns("StoreID").DataField = "StoreID"
         GridStore.Columns("Config").DataField = "Config"
         GridStore.MoveFirst
         For i = 1 To GridStore.Rows
               vStoreID = GridStore.Columns("StoreID").Value
               vConfig = GridStore.Columns("Config").Value
               vAccountNo = GridStore.Columns("AccountNo").Value
               Call StockTransferExport
               If i < GridStore.Rows Then GridStore.MoveNext
         Next i
         
         MsgBox "Sync Succeed", vbInformation, Me.Caption
         Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub BtnTransactionImport_Click()
   On Error GoTo ErrorHandler
      Me.MousePointer = vbHourglass
      vSQL = "Select * from Stores where StoreID <> 1 and isLock = 0 and Config is not null"
      If Rs.State = adStateOpen Then Rs.Close
      Rs.Open vSQL, CN, adOpenStatic, adLockReadOnly
   
      While Not Rs.EOF
         vStoreID = Rs!StoreID
         vConfig = Rs!Config
         Call TransactionImport
         Rs.MoveNext
      Wend
         MsgBox "Sync Succeed", vbInformation, Me.Caption
         Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub DefinationExport()
On Error GoTo ErrorHandler
For vTableCount = 1 To LVDef.ListItems.Count
   If LVDef.ListItems(vTableCount).Checked = True Then
      
      GetTableColumn (LVDef.ListItems(vTableCount).Text)
      
      ''''''' Make Update Query
      vSQL = "Update " & vConfig & LVDef.ListItems(vTableCount).Text & " Set "
      GridTable.MoveFirst
      For i = 1 To GridTable.Rows
         If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("Tag") Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID") Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
               vSQL = vSQL & vbCrLf _
               & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
         End If
         GridTable.MoveNext
      Next i
      GridTable.MoveFirst
      
      vSQL = Replace(vSQL, "''", "Null")
      vSQL = Left(vSQL, Len(vSQL) - 2)
      
      If LVDef.ListItems(vTableCount).Text = "ChartOfAccounts" Then
         vSQL = vSQL & vbCrLf _
         & "From " & vConfig & LVDef.ListItems(vTableCount).Text & " t1 " & vbCrLf _
         & "Inner Join " & vbCrLf _
         & "(" & vbCrLf _
         & " Select C.*" & vbCrLf _
         & " from " & LVDef.ListItems(vTableCount).Text & " C" & vbCrLf _
         & " Left Outer Join Parties P on P.partyID =  C.AccountNo " & vbCrLf _
         & " Left Outer Join Employees  E on E.EmpID =  C.AccountNo  " & vbCrLf _
         & " Where C.IsSync = 0 And C.Accountno Like ('6%') And (P.StoreID is null or P.StoreID = " & vStoreID & ") And (P.StoreID is null or P.StoreID = " & vStoreID & ")" & vbCrLf _
         & " )T2"
         vSQL = vSQL & vJoin
         CN.Execute vSQL
         
      Else
         vSQL = vSQL & vbCrLf _
         & "From " & vConfig & LVDef.ListItems(vTableCount).Text & " t1 " & vbCrLf _
         & "Inner Join " & vbCrLf _
         & "(" & vbCrLf _
         & " Select *" & vbCrLf _
         & " from " & LVDef.ListItems(vTableCount).Text & vbCrLf _
         & " Where IsSync = 0 " & vbCrLf _
         & " )T2"
         vSQL = vSQL & vJoin
         CN.Execute vSQL
      End If
      
      
      ''''''' Insert Query
      If LVDef.ListItems(vTableCount).Text = "ChartOfAccounts" Then
         
         vSQL = "Insert into " & vConfig & LVDef.ListItems(vTableCount).Text & IIf(vIdentityKey = 1, "(" & vColumnList1 & ")", "") & vbCrLf _
               & " Select " & IIf(vIdentityKey = 1, vColumnList1, "T1.*") & " from " & LVDef.ListItems(vTableCount).Text & " T1 Left outer join  " & vConfig & LVDef.ListItems(vTableCount).Text & " T2" & vbCrLf _
               & vJoin & vbCrLf _
               & "Left Outer Join Parties P on P.partyID =  T1.AccountNo "
         vSQL = vSQL & vbCrLf _
            & "where T2." & vPKey1 & "  is null And T1.Accountno Like ('6%') And (P.StoreID is null or P.StoreID = " & vStoreID & ")"
   
         CN.Execute vSQL
      ElseIf LVDef.ListItems(vTableCount).Text = "Parties" Then
         vSQL = "Insert into " & vConfig & LVDef.ListItems(vTableCount).Text & IIf(vIdentityKey = 1, "(" & vColumnList1 & ")", "") & vbCrLf _
               & " Select " & IIf(vIdentityKey = 1, vColumnList1, "T1.*") & " from " & LVDef.ListItems(vTableCount).Text & " T1 Left outer join  " & vConfig & LVDef.ListItems(vTableCount).Text & " T2"
         vSQL = vSQL & vbCrLf _
            & vJoin & vbCrLf _
            & "where T2." & vPKey1 & "  is null And (T1.StoreID is null or T1.StoreID = " & vStoreID & ")"
   
         CN.Execute vSQL
      ElseIf LVDef.ListItems(vTableCount).Text = "Employees" Then
         vSQL = "Insert into " & vConfig & LVDef.ListItems(vTableCount).Text & IIf(vIdentityKey = 1, "(" & vColumnList1 & ")", "") & vbCrLf _
               & " Select " & IIf(vIdentityKey = 1, vColumnList1, "T1.*") & " from " & LVDef.ListItems(vTableCount).Text & " T1 Left outer join  " & vConfig & LVDef.ListItems(vTableCount).Text & " T2"
         vSQL = vSQL & vbCrLf _
            & vJoin & vbCrLf _
            & "where T2." & vPKey1 & "  is null And (T1.StoreID is null or T1.StoreID = " & vStoreID & ")"
   
         CN.Execute vSQL
      ElseIf LVDef.ListItems(vTableCount).Text = "Users" Then
         vSQL = "Insert into " & vConfig & LVDef.ListItems(vTableCount).Text & IIf(vIdentityKey = 1, "(" & vColumnList1 & ")", "") & vbCrLf _
               & " Select " & IIf(vIdentityKey = 1, vColumnList1, "T1.*") & " from " & LVDef.ListItems(vTableCount).Text & " T1 Left outer join  " & vConfig & LVDef.ListItems(vTableCount).Text & " T2"
         vSQL = vSQL & vbCrLf _
            & vJoin & vbCrLf _
            & "where T2." & vPKey1 & "  is null And (T1.StoreID is null or T1.StoreID = " & vStoreID & ")"
   
         CN.Execute vSQL
      Else
         vSQL = "Insert into " & vConfig & LVDef.ListItems(vTableCount).Text & IIf(vIdentityKey = 1, "(" & vColumnList1 & ")", "") & vbCrLf _
               & " Select " & IIf(vIdentityKey = 1, vColumnList1, "T1.*") & " from " & LVDef.ListItems(vTableCount).Text & " T1 Left outer join  " & vConfig & LVDef.ListItems(vTableCount).Text & " T2"
         vSQL = vSQL & vbCrLf _
            & vJoin & vbCrLf _
            & "where T2." & vPKey1 & "  is null" '" Where T2.modified_on >=  (Select isnull(max(modified_on),'01-01-1900')  modified_on from " & LVDef.ListItems(vTableCount).Text & " )" & vbCrLf _

         CN.Execute vSQL
         
      End If
            
   End If
     
Next vTableCount
Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   MsgBox Err.Description
   Call ShowErrorMessage
End Sub

Private Sub SyncDefination2()

For vTableCount = 1 To LVDef.ListItems.Count
   If LVDef.ListItems(vTableCount).Checked = True Then
   vIdentityKey = 0
   vPKey1 = ""
   vPKey2 = ""
   vPKey3 = ""
   
   vHeaderTable = LVDef.ListItems(vTableCount).SubItems(1)
   
   ''''''' Get Primary Key Column of table
   vSQL = "SELECT Column_Name From INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE OBJECTPROPERTY(OBJECT_ID(CONSTRAINT_SCHEMA + '.' + QUOTENAME(CONSTRAINT_NAME)), 'IsPrimaryKey') = 1 AND TABLE_NAME = '" & vHeaderTable & "' AND TABLE_SCHEMA = 'dbo'"
   With CN.Execute(vSQL)
      For vPKeyCount = 1 To .RecordCount
         Select Case vPKeyCount
               Case 1
                  vPKey1 = .Fields(0)
                  vJoin = " on T2." & vPKey1 & " = T1." & vPKey1
               Case 2
                  vPKey2 = .Fields(0)
                  vJoin = " on T2." & vPKey1 & " = T1." & vPKey1 & " and T2." & vPKey2 & " = T1." & vPKey2
               Case 3
                  vPKey3 = .Fields(0)
                  vJoin = " on T2." & vPKey1 & " = T1." & vPKey1 & " and T2." & vPKey2 & " = T1." & vPKey2 & " and T2." & vPKey3 & " = T1." & vPKey3
         End Select
         .MoveNext
      Next vPKeyCount
      vPKeyCount = vPKeyCount - 1
   End With
   
   '''''''' Get Column of table
   vSQL = "Select Column_Name, Data_Type FROM INFORMATION_SCHEMA.COLUMNS Where TABLE_NAME = N'" & LVDef.ListItems(vTableCount).Text & "'"
   With CN.Execute(vSQL)
      GridTable.MoveFirst
      GridTable.RemoveAll
      GridTable.AllowAddNew = True
      While Not .EOF
         GridTable.AddNew
         GridTable.Columns("Column_Name").Text = !Column_Name
         GridTable.Columns("Data_Type").Text = !Data_Type
         .MoveNext
      Wend
      GridTable.MoveFirst
      .Close
   End With
   
   ''''''' Get Identity Key Column of table
   vSQL = "SELECT Column_Name From INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE COLUMNPROPERTY(object_id(TABLE_NAME), COLUMN_NAME, 'IsIdentity') = 1 AND TABLE_NAME = '" & LVDef.ListItems(vTableCount).Text & "' AND TABLE_SCHEMA = 'dbo'"
   With CN.Execute(vSQL)
         If .RecordCount > 0 Then
            vIdentityKey = 1
            vSerialNo = .Fields(0)
'            vPKeyCount = vPKeyCount + vIdentityKey
            GridTable.Row = 1
            vColumnList1 = ""
            vColumnList2 = ""
            For i = 1 To GridTable.Rows - 1
               vColumnList1 = vColumnList1 + "T2." + GridTable.Columns("Column_Name").Value & ", "
               vColumnList2 = vColumnList2 + GridTable.Columns("Column_Name").Value & ", "
               GridTable.MoveNext
            Next i
            vColumnList1 = Left(vColumnList1, Len(vColumnList1) - 2)
            vColumnList2 = Left(vColumnList2, Len(vColumnList2) - 2)
            GridTable.MoveFirst
         End If
   End With
   
   '   vSQL = "Select * from " & vLinkedServer & LVDef.ListItems(vTableCount).Text & vbCrLf _
         & " Where modified_on >  (Select max(modified_on)  modified_on from " & LVDef.ListItems(vTableCount).Text & " )"
   
   
   ''''''' Make Update Query
   vSQL = "Update " & LVDef.ListItems(vTableCount).Text & " Set "
   GridTable.MoveFirst
   For i = 1 To GridTable.Rows
      If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
         vSQL = vSQL & vbCrLf _
         & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
      End If
      GridTable.MoveNext
   Next i
   GridTable.MoveFirst
   vSQL = Replace(vSQL, "''", "Null")
   vSQL = Left(vSQL, Len(vSQL) - 2)
   vSQL = vSQL & vbCrLf _
   & "From " & LVDef.ListItems(vTableCount).Text & " t1 " & vbCrLf _
   & "Inner Join " & vbCrLf _
   & "(" & vbCrLf _
   & " Select * from " & vLinkedServer & LVDef.ListItems(vTableCount).Text & vbCrLf _
   & " Where modified_on >  (Select max(modified_on)  modified_on from " & LVDef.ListItems(vTableCount).Text & " )" & vbCrLf _
   & " )T2"
   vSQL = vSQL & vJoin
   CN.Execute vSQL
   
         
   ''''''' Insert record
      If LVDef.ListItems(vTableCount).Text <> "Products" Then
         vSQL = "Insert into " & LVDef.ListItems(vTableCount).Text & IIf(vIdentityKey = 1, "(" & vColumnList2 & ")", "") & vbCrLf _
               & " Select " & IIf(vIdentityKey = 1, vColumnList1, "T2.*") & " from " & LVDef.ListItems(vTableCount).Text & " T1 right outer join  " & vLinkedServer & LVDef.ListItems(vTableCount).Text & " T2"
         vSQL = vSQL & vbCrLf _
            & vJoin & vbCrLf _
            & "where T1." & vPKey1 & "  is null" '" Where T2.modified_on >=  (Select isnull(max(modified_on),'01-01-1900')  modified_on from " & LVDef.ListItems(vTableCount).Text & " )" & vbCrLf _

         CN.Execute vSQL
      Else
         vSQL = "Select T2.* from " & LVDef.ListItems(vTableCount).Text & " T1 right outer join  " & vLinkedServer & LVDef.ListItems(vTableCount).Text & " T2 " & vbCrLf _
         & vJoin & vbCrLf _
         & "where T1." & vPKey1 & "  is null"
         With CN.Execute(vSQL)
            While Not .EOF
               vSQL = "Insert into " & LVDef.ListItems(vTableCount).Text & IIf(vIdentityKey = 1, "(" & vColumnList2 & ")", "") & vbCrLf _
               & " Select " & IIf(vIdentityKey = 1, vColumnList1, "T2.*") & " from " & LVDef.ListItems(vTableCount).Text & " T1 right outer join  " & vLinkedServer & LVDef.ListItems(vTableCount).Text & " T2"
                  vSQL = vSQL & vbCrLf _
                  & vJoin
               Select Case vPKeyCount
                  Case 1
                  vwhere = " Where T2." & vPKey1 & " = '" & .Fields(vPKey1).Value & "'"
               End Select
               vSQL = vSQL & vwhere
               CN.Execute vSQL
               .MoveNext
            Wend
         End With
      End If
      LVDef.ListItems(vTableCount).Checked = False
   End If
   
   
   Next vTableCount

End Sub



Private Sub SyncStockTransfer()
   On Error GoTo ErrorHandler
   Dim i As Integer, j As Integer
   Dim vMainStoreID As Byte
  
   
   vMainStoreID = CN.Execute("Select StoreID from Stores").Fields(0).Value
   
   
      ''''''''''' Update Purchase Pending Header
      vSQL = "Update PurchasePendingHeader Set " & vbCrLf _
            & "UserNo = T2.UserNo " & vbCrLf _
            & "From PurchasePendingHeader T1 inner join " & vbCrLf _
            & "(Select ToStoreID, UserNo from " & vLinkedServer & "StockTransferHeader where ToStoreID = " & vMainStoreID & vbCrLf _
            & "And modified_on >  (Select max(modified_on)  modified_on from " & vLinkedServer & " StockTransferHeader Where ToStoreID = " & vMainStoreID & " )) T2 on T2.toStoreID = T1.StoreID "
'      CN.Execute vSQL

      ''''''''''' Update Purchase Pending Body
      vSQL = "Update PurchasePendingBody Set " & vbCrLf _
            & "Code = T2.Code, ProductID = T2.ProductID, QtyPack = T2.QtyPack, QtyLoose = T2.QtyLoose " & vbCrLf _
            & "From PurchasePendingBody T1 inner join " & vbCrLf _
            & "(Select ToStoreID, UserNo from " & vLinkedServer & "StockTransferBody where ToStoreID = " & vMainStoreID & vbCrLf _
            & "And modified_on >  (Select max(modified_on)  modified_on from " & vLinkedServer & " StockTransferHeader Where ToStoreID = " & vMainStoreID & " )) T2 on T2.toStoreID = T1.StoreID "
'      CN.Execute vSQL
      
      ''''''''''' Insert Purchase Pending Header
       vSQL = "Select T2.* from PurchasePendingHeader T1 right outer join  " & vLinkedServer & "StockTransferHeader T2 " & vbCrLf _
         & "On T1.storeID = T2.ToStoreID" & vbCrLf _
         & "where T1.StoreID is null"
         With CN.Execute(vSQL)
            While Not .EOF
               
               FunGetMaxID = CN.Execute("Select isnull(max(PurID),0)+1 from PurchasePendingHeader").Fields(0).Value
               vSQL = "Insert into PurchasePendingHeader (PurID, PurchaseDate, VendorID, TotalAmount, StoreID, UserNo ) Values " & vbCrLf _
                      & "(" & FunGetMaxID & ",'" & .Fields("TransferDate").Value & "',611,0," & .Fields("ToStoreID").Value & "," & .Fields("UserNo").Value & ")"
               CN.Execute vSQL
               
               ''''''''''' Insert Purchase Pending Body
               vSQL = "Select T2.* from PurchasePendingBody T1 right outer join  " & vLinkedServer & "StockTransferBody T2 " & vbCrLf _
               & "On T1.storeID = T2.StoreID and T2.TransferID = " & .Fields("TransferID").Value & " and T2.TransferDate = '" & .Fields("TransferDate").Value & "'" & vbCrLf _
               & "where T1.StoreID is null"
               If Rs.State = adStateOpen Then Rs.Close
               Rs.Open vSQL, CN, adOpenStatic, adLockBatchOptimistic
                  vTotalAmount = 0
                  While Not Rs.EOF
                     vPrice = CN.Execute("Select Isnull(PurPrice,0) PurPrice from products where productId = '" & Rs!ProductID & "'").Fields(0).Value
                     vMultiplier = 0
                        With CN.Execute("Select isnull(Multiplier,0) Multiplier from productPacking where productId = '" & Rs!ProductID & "'")
                           If .RecordCount > 0 Then vMultiplier = .Fields(0).Value
                        End With
                     vAmount = Round(vPrice * ((IIf(IsNull(Rs!QtyPack), 0, Rs!QtyPack) * vMultiplier) + Rs!QtyLoose), 2)
                     vTotalAmount = vTotalAmount + vAmount
                     vSQL = "Insert into PurchasePendingBody (PurID, PurchaseDate,  Code, ProductID, QtyPack, QtyLoose, StoreID, Price, DiscVal, DiscPer, DiscPC, Amount ) Values " & vbCrLf _
                           & "(" & FunGetMaxID & ",'" & Rs!TransferDate & "'," & Rs!Code & ",'" & Rs!ProductID & "'," & Rs!QtyPack & "," & Rs!QtyLoose & "," & Rs!QtyLoose & ",0,0,0,0," & vAmount & ")"
                     CN.Execute vSQL
                     CN.Execute ("Update PurchasePendingHeader Set TotalAmount = " & vTotalAmount & " Where PurID = " & FunGetMaxID & " And PurchaseDate = '" & Rs!TransferDate & "'")
                     Rs.MoveNext
                  Wend
               .MoveNext
            Wend
         End With
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
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

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hwnd, "Sync Data"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   CN.CommandTimeout = 0
   Call Settings
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Settings()
   On Error GoTo ErrorHandler
      GridTable.CancelUpdate
      GridTable.RemoveAll
      GridTable.AddNew
      GridTable.Columns("Column_Name").Text = " "
      GridTable.Update
      
      
      Call SetDefinations
      
      Call SetTransactions
      
     
      
            
      
      
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SetDefinations()
      
      LVDef.FullRowSelect = True
      LVDef.ListItems.Clear
      LVDef.ColumnHeaders.Add , , "Serial", 50, 0
      LVDef.ColumnHeaders.Add , , "Name", 250, 0
      LVDef.View = lvwReport
      
      Set Item = LVDef.ListItems.Add(, , "Companies")
      Item.SubItems(1) = "Companies"
      Item.Checked = True
      
      Set Item = LVDef.ListItems.Add(, , "Groups")
      Item.SubItems(1) = "Groups"
      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "SubGroups")
      Item.SubItems(1) = "SubGroups"
      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "Brands")
      Item.SubItems(1) = "Brands"
      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "Departments")
      Item.SubItems(1) = "Departments"
      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "SubDepartments")
      Item.SubItems(1) = "SubDepartments"
      Item.Checked = False

      Set Item = LVDef.ListItems.Add(, , "Colours")
      Item.SubItems(1) = "Colours"
      Item.Checked = False

      Set Item = LVDef.ListItems.Add(, , "Sizes")
      Item.SubItems(1) = "Sizes"
      Item.Checked = False

      Set Item = LVDef.ListItems.Add(, , "Seasons")
      Item.SubItems(1) = "Seasons"
      Item.Checked = False

      Set Item = LVDef.ListItems.Add(, , "Descriptions")
      Item.SubItems(1) = "Descriptions"
      Item.Checked = False

      Set Item = LVDef.ListItems.Add(, , "ItemDescription")
      Item.SubItems(1) = "ItemDescription"
      Item.Checked = False

      Set Item = LVDef.ListItems.Add(, , "Parties")
      Item.SubItems(1) = "Parties "
      Item.Checked = False
      
      Set Item = LVDef.ListItems.Add(, , "Users")
      Item.SubItems(1) = "Users"
      Item.Checked = False

      Set Item = LVDef.ListItems.Add(, , "Products")
      Item.SubItems(1) = "Products"
      Item.Checked = True
                          
      Set Item = LVDef.ListItems.Add(, , "ChartOfAccounts")
      Item.SubItems(1) = "ChartOfAccounts"
      Item.Checked = False

      Set Item = LVDef.ListItems.Add(, , "Designations")
      Item.SubItems(1) = "Designations"
      Item.Checked = False

      Set Item = LVDef.ListItems.Add(, , "EmpDepartments")
      Item.SubItems(1) = "EmpDepartments"
      Item.Checked = False

      Set Item = LVDef.ListItems.Add(, , "Employees")
      Item.SubItems(1) = "Employees"
      Item.Checked = False
         
      Set Item = LVDef.ListItems.Add(, , "Members")
      Item.SubItems(1) = "Members"
      Item.Checked = False
      
      Set Item = LVDef.ListItems.Add(, , "MemberTypes")
      Item.SubItems(1) = "MemberTypes"
      Item.Checked = False
      
      Set Item = LVDef.ListItems.Add(, , "MembersDiscount")
      Item.SubItems(1) = "MembersDiscount"
      Item.Checked = False
      
      Set Item = LVDef.ListItems.Add(, , "ProductBarcodes")
      Item.SubItems(1) = "ProductBarcodes"
      Item.Checked = True
     
     



End Sub

Private Sub SetTransactions()
      LVTrans.FullRowSelect = True
      LVTrans.ListItems.Clear
      LVTrans.ColumnHeaders.Add , , "Table", 150, 0
      LVTrans.ColumnHeaders.Add , , "Header", 150, 0
      
      LVTrans.View = lvwReport
      
      Set Item = LVTrans.ListItems.Add(, , "Sale")
      Item.SubItems(1) = "SaleHeader"
      Item.Checked = False


      Set Item = LVTrans.ListItems.Add(, , "SaleReturn")
      Item.SubItems(1) = "SaleReturn"
      Item.Checked = False
      
      Set Item = LVTrans.ListItems.Add(, , "Purchase")
      Item.SubItems(1) = "PurchaseHeader"
      Item.Checked = False
      
      Set Item = LVTrans.ListItems.Add(, , "PurchaseReturn")
      Item.SubItems(1) = "PurchaseReturn"
      Item.Checked = False
      
      Set Item = LVTrans.ListItems.Add(, , "CreditVouchers")
      Item.SubItems(1) = "CreditVouchers"
      Item.Checked = False

      Set Item = LVTrans.ListItems.Add(, , "DebitVouchers")
      Item.SubItems(1) = "DebitVouchers"
      Item.Checked = False


      Set Item = LVTrans.ListItems.Add(, , "JournalVouchers")
      Item.SubItems(1) = "JournalVouchers"
      Item.Checked = False
      
'      Set Item = LVTrans.ListItems.Add(, , "UserClosingHeader")
'      Item.SubItems(1) = "UserClosingHeader"
'      Item.Checked = False
'
'      Set Item = LVTrans.ListItems.Add(, , "UserClosingBody")
'      Item.SubItems(1) = "UserClosingHeader"
'      Item.Checked = False
'
'      Set Item = LVTrans.ListItems.Add(, , "AdminClosing")
'      Item.SubItems(1) = "AdminClosing"
'      Item.Checked = False
            
End Sub



Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub Timer1_Timer()
   On Error GoTo ErrorHandler
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
Private Sub TransactionImport()
   
      For vTableCount = 1 To LVTrans.ListItems.Count
      
         If LVTrans.ListItems(vTableCount).Checked = True Then
              
            vHeaderTable = LVTrans.ListItems(vTableCount).SubItems(1)
            
            Select Case LVTrans.ListItems(vTableCount).Text
               Case "Sale"
                  Call SaleTransaction
               Case "SaleReturn"
                  Call SaleReturnTransaction
               Case "Purchase"
                  Call PurchaseTransaction
               Case "CreditVouchers"
                  Call CreditVouchersTransaction
               Case "DebitVouchers"
                  Call DebitVouchersTransaction
               Case "JournalVouchers"
'                  Call JournalVouchersTransaction
            End Select
   
'         LVTrans.ListItems(vTableCount).Checked = False
      End If
      Next vTableCount
     
End Sub
Private Sub GetTableColumn(vTable As String)
            
   
   ''''''' Get Primary Key Column of table
   vSQL = "SELECT Column_Name From INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE OBJECTPROPERTY(OBJECT_ID(CONSTRAINT_SCHEMA + '.' + QUOTENAME(CONSTRAINT_NAME)), 'IsPrimaryKey') = 1 AND TABLE_NAME = '" & vTable & "' AND TABLE_SCHEMA = 'dbo'"
   With CN.Execute(vSQL)
      For vPKeyCount = 1 To .RecordCount
         Select Case vPKeyCount
               Case 1
                  vPKey1 = .Fields(0)
                  vJoin = " on T2." & vPKey1 & " = T1." & vPKey1
               Case 2
                  vPKey2 = .Fields(0)
                  vJoin = " on T2." & vPKey1 & " = T1." & vPKey1 & " and T2." & vPKey2 & " = T1." & vPKey2
               Case 3
                  vPKey3 = .Fields(0)
                  vJoin = " on T2." & vPKey1 & " = T1." & vPKey1 & " and T2." & vPKey2 & " = T1." & vPKey2 & " and T2." & vPKey3 & " = T1." & vPKey3
         End Select
         .MoveNext
      Next vPKeyCount
      vPKeyCount = vPKeyCount - 1
   End With
   '''''''' Get Column of table
   vSQL = "Select Column_Name, Data_Type FROM INFORMATION_SCHEMA.COLUMNS Where TABLE_NAME = N'" & vTable & "'"
   With CN.Execute(vSQL)
      GridTable.MoveFirst
      GridTable.RemoveAll
      GridTable.AllowAddNew = True
      While Not .EOF
         GridTable.AddNew
         GridTable.Columns("Column_Name").Text = !Column_Name
         GridTable.Columns("Data_Type").Text = !Data_Type
         .MoveNext
      Wend
      GridTable.MoveFirst
      .Close
   End With
   vColumnList1 = ""
   vColumnList2 = ""
   ''''''' Get Identity Key Column of table
   vSQL = "SELECT Column_Name From INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE COLUMNPROPERTY(object_id(TABLE_NAME), COLUMN_NAME, 'IsIdentity') = 1 AND TABLE_NAME = '" & vTable & "' AND TABLE_SCHEMA = 'dbo'"
   With CN.Execute(vSQL)
         If .RecordCount > 0 Then
            vIdentityKey = 1
            vSerialNo = .Fields(0)
'            vPKeyCount = vPKeyCount + vIdentityKey
            GridTable.Row = 0
            vColumnList1 = ""
            vColumnList2 = ""
            vColumnList2 = ""
            For i = 0 To GridTable.Rows - 1
               If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID")) Then
                  vColumnList1 = vColumnList1 + GridTable.Columns("Column_Name").Value & ", "
                  vColumnList2 = vColumnList2 + "T2." + GridTable.Columns("Column_Name").Value & ", "
               End If
               GridTable.MoveNext
            Next i
            vColumnList1 = Left(vColumnList1, Len(vColumnList1) - 2)
            vColumnList2 = Left(vColumnList2, Len(vColumnList2) - 2)
            GridTable.MoveFirst
         End If
   End With

End Sub


Private Sub SaleTransaction()
On Error GoTo ErrorHandler
   GetTableColumn ("SaleHeader")
   
   ''''''' Make Update Query
   vSQL = "Update SaleHeader Set "
   GridTable.MoveFirst
   For i = 1 To GridTable.Rows
      If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("Tag") Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID") Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
         If UCase(GridTable.Columns("Column_Name").Value) = UCase("StoreID") Then
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = " & vStoreID & ", "
         Else
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
         End If
      End If
      GridTable.MoveNext
   Next i
   GridTable.MoveFirst
   
   vSQL = Replace(vSQL, "''", "Null")
   vSQL = Left(vSQL, Len(vSQL) - 2)
   vSQL = vSQL & vbCrLf _
   & "From SaleHeader t1 " & vbCrLf _
   & "Inner Join " & vbCrLf _
   & "(" & vbCrLf _
   & " Select SID, " & vColumnList1 & vbCrLf _
   & " from " & vConfig & "SaleHeader " & vbCrLf _
   & " Where IsSync = 0 And Tag is null" & vbCrLf _
   & " )T2"
   vSQL = vSQL & vbCrLf _
   & " on T1.Tag = T2.SID And T1.BillDate = T2.BillDate And T1.StoreID = " & vStoreID & vbCrLf _
   & " Where T1.Tag is not Null "
   CN.Execute vSQL
         
   
   
   ''''''' Insert record
   vColumnList2 = Replace(UCase(vColumnList2), "TAG", "SID")
   vColumnList2 = Replace(UCase(vColumnList2), UCase("T2.StoreID"), vStoreID)
   
   vSQL = "Insert into " & " SaleHeader " & "(" & vColumnList1 & ")" & vbCrLf _
         & " Select " & vColumnList2 & vbCrLf _
         & " from " & vConfig & "SaleHeader T2" & vbCrLf _
         & " Left outer join  " & " SaleHeader T1 "
   vSQL = vSQL & vbCrLf _
         & " on T1.Tag = T2.SID And T1.BillDate = T2.BillDate And T1.StoreID = " & vStoreID & vbCrLf _
         & " Where T1.Tag is Null And T2.IsSync = 0 "
    CN.Execute vSQL
   
  
         
   GetTableColumn ("SaleBody")
   
    ''''''' Make Update Query
   vSQL = "Update SaleBody Set "
   
   GridTable.MoveFirst
   For i = 1 To GridTable.Rows
      If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("SID") Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID") Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
         If UCase(GridTable.Columns("Column_Name").Value) = UCase("StoreID") Then
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = " & vStoreID & ", "
         Else
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
         End If
      End If
      GridTable.MoveNext
   Next i
   GridTable.MoveFirst
   
   vSQL = Replace(vSQL, "''", "Null")
   vSQL = Left(vSQL, Len(vSQL) - 2)
   vSQL = vSQL & vbCrLf _
   & "From SaleBody B " & vbCrLf _
   & "inner join SaleHeader H on H.SID = B.SID" & vbCrLf _
   & "Inner Join " & vbCrLf _
   & "(" & vbCrLf _
   & " Select " & vColumnList2 & vbCrLf _
   & " from " & vConfig & "SaleBody T2 " & vbCrLf _
   & " inner join " & vConfig & "SaleHeader T1 " & vbCrLf _
   & " on T1.SID = T2.SID And T1.BillDate = T2.BillDate " & vbCrLf _
   & " inner join  SaleHeader H on H.Tag = T1.SID And H.StoreID = " & vStoreID & vbCrLf _
   & " Where T1.Tag is Null And T1.IsSync = 0 And H.StoreID = " & vStoreID & vbCrLf _
   & " )T2"
   vSQL = vSQL & vbCrLf _
   & " on H.Tag = T2.SID And H.BillDate = T2.BillDate And T2.ProductID = B.ProductID And H.StoreID = " & vStoreID & vbCrLf _
   & " Where H.Tag is not Null "
   CN.Execute vSQL

   ''''''' Insert record
          
   vColumnList2 = Replace(UCase(vColumnList2), "T2.SID", "H.SID")
   vColumnList2 = Replace(UCase(vColumnList2), UCase("T2.StoreID"), vStoreID)
   
   vSQL = "Insert into " & " SaleBody " & "(" & vColumnList1 & ")" & vbCrLf _
         & " Select " & vColumnList2 & vbCrLf _
         & " from " & vConfig & "SaleBody T2" & vbCrLf _
         & " inner join " & vConfig & "SaleHeader T1" & vbCrLf _
         & " on T2.SID = T1.SID And T2.BillDate = T1.BillDate" & vbCrLf _
         & " inner join  SaleHeader H on H.Tag = T1.SID And H.StoreID =" & vStoreID & vbCrLf _
         & " left outer Join SaleBody B on B.SID = H.SID and B.BillID = T2.BillID and B.BILLDATE = T2.BillDate And B.ProductID = T2.ProductID" & vbCrLf _
         & " Where T1.IsSync = 0 And B.SerialNo is Null And T1.Tag is Null And H.StoreID =" & vStoreID
      
   CN.Execute vSQL
     
   ''''''''''''' Update Remotly Client System '''''''''''''''''''''
   vSQL = "Update " & vConfig & "SaleHeader Set IsSync =1 where IsSync = 0"
   CN.Execute vSQL
   
   ''''''''''''' Delete Sale Body '''''''''''''''''''''
   vSQL = "Delete SaleBody " & vbCrLf _
          & " From SaleBody b" & vbCrLf _
          & " Inner join SaleHeader H on H.SID = B.SID And H.BillDate = B.BillDate" & vbCrLf _
          & " Left outer join " & vConfig & "SaleHeader T1 on T1.SID = H.Tag And H.StoreID = " & vStoreID & vbCrLf _
          & " Left outer join " & vConfig & "SaleBody T2 on T1.SID = T2.SID And T1.SID = H.Tag And T1.BillDate = T2.BillDate And T2.ProductID = b.ProductID" & vbCrLf _
          & " WHere H.Tag is Not Null And T2.SerialNo is null And H.StoreID = " & vStoreID
  CN.Execute vSQL
  ''''''''''''' Delete Sale SaleHeader '''''''''''''''''''''
  vSQL = "Delete SaleHeader " & vbCrLf _
          & " From SaleHeader H" & vbCrLf _
          & " Left outer join " & vConfig & "SaleHeader T1 on T1.SID = H.Tag And H.StoreID = " & vStoreID & vbCrLf _
          & " WHere H.Tag is Not Null And T1.SID is null And H.StoreID = " & vStoreID
   CN.Execute vSQL
Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
   
End Sub

Private Sub SaleReturnTransaction()
On Error GoTo ErrorHandler
   GetTableColumn ("SaleReturnHeader")
   
   ''''''' Make Update Query
   vSQL = "Update SaleReturnHeader Set "
   GridTable.MoveFirst
   For i = 1 To GridTable.Rows
      If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("Tag") Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID") Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
         If UCase(GridTable.Columns("Column_Name").Value) = UCase("StoreID") Then
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = " & vStoreID & ", "
         Else
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
         End If
      End If
      GridTable.MoveNext
   Next i
   GridTable.MoveFirst
   
   vSQL = Replace(vSQL, "''", "Null")
   vSQL = Left(vSQL, Len(vSQL) - 2)
   vSQL = vSQL & vbCrLf _
   & "From SaleReturnHeader t1 " & vbCrLf _
   & "Inner Join " & vbCrLf _
   & "(" & vbCrLf _
   & " Select SID, " & vColumnList1 & vbCrLf _
   & " from " & vConfig & "SaleReturnHeader " & vbCrLf _
   & " Where IsSync = 0 And Tag is null" & vbCrLf _
   & " )T2"
   vSQL = vSQL & vbCrLf _
   & " on T1.Tag = T2.SID And T1.ReturnDate = T2.ReturnDate And T1.StoreID = " & vStoreID & vbCrLf _
   & " Where T1.Tag is not Null "
   CN.Execute vSQL
         
   
   
   ''''''' Insert record
   vColumnList2 = Replace(UCase(vColumnList2), "TAG", "SID")
   vColumnList2 = Replace(UCase(vColumnList2), UCase("T2.StoreID"), vStoreID)
   
   vSQL = "Insert into " & " SaleReturnHeader " & "(" & vColumnList1 & ")" & vbCrLf _
         & " Select " & vColumnList2 & vbCrLf _
         & " from " & vConfig & "SaleReturnHeader T2" & vbCrLf _
         & " Left outer join  " & " SaleReturnHeader T1 "
   vSQL = vSQL & vbCrLf _
         & " on T1.Tag = T2.SID And T1.ReturnDate = T2.ReturnDate And T1.StoreID = " & vStoreID & vbCrLf _
         & " Where T1.Tag is Null And T2.IsSync = 0 "
    CN.Execute vSQL
   
  
         
   GetTableColumn ("SaleReturnBody")
   
    ''''''' Make Update Query
   vSQL = "Update SaleReturnBody Set "
   
   GridTable.MoveFirst
   For i = 1 To GridTable.Rows
      If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("SID") Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID") Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
         If UCase(GridTable.Columns("Column_Name").Value) = UCase("StoreID") Then
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = " & vStoreID & ", "
         Else
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
         End If
      End If
      GridTable.MoveNext
   Next i
   GridTable.MoveFirst
   
   vSQL = Replace(vSQL, "''", "Null")
   vSQL = Left(vSQL, Len(vSQL) - 2)
   vSQL = vSQL & vbCrLf _
   & "From SaleReturnBody B " & vbCrLf _
   & "inner join SaleReturnHeader H on H.SID = B.SID" & vbCrLf _
   & "Inner Join " & vbCrLf _
   & "(" & vbCrLf _
   & " Select " & vColumnList2 & vbCrLf _
   & " from " & vConfig & "SaleReturnBody T2 " & vbCrLf _
   & " inner join " & vConfig & "SaleReturnHeader T1 " & vbCrLf _
   & " on T1.SID = T2.SID And T1.ReturnDate = T2.ReturnDate " & vbCrLf _
   & " inner join  SaleReturnHeader H on H.Tag = T1.SID And H.StoreID = " & vStoreID & vbCrLf _
   & " Where T1.Tag is Null And T1.IsSync = 0 And H.StoreID = " & vStoreID & vbCrLf _
   & " )T2"
   vSQL = vSQL & vbCrLf _
   & " on H.Tag = T2.SID And H.ReturnDate = T2.ReturnDate And T2.ProductID = B.ProductID And H.StoreID = " & vStoreID & vbCrLf _
   & " Where H.Tag is not Null "
   CN.Execute vSQL

   ''''''' Insert record
          
   vColumnList2 = Replace(UCase(vColumnList2), "T2.SID", "H.SID")
   vColumnList2 = Replace(UCase(vColumnList2), UCase("T2.StoreID"), vStoreID)
   
   vSQL = "Insert into " & " SaleReturnBody " & "(" & vColumnList1 & ")" & vbCrLf _
         & " Select " & vColumnList2 & vbCrLf _
         & " from " & vConfig & "SaleReturnBody T2" & vbCrLf _
         & " inner join " & vConfig & "SaleReturnHeader T1" & vbCrLf _
         & " on T2.SID = T1.SID And T2.ReturnDate = T1.ReturnDate" & vbCrLf _
         & " inner join  SaleReturnHeader H on H.Tag = T1.SID And H.StoreID =" & vStoreID & vbCrLf _
         & " left outer Join SaleReturnBody B on B.SID = H.SID and B.ReturnID = T2.ReturnID and B.ReturnDate = T2.ReturnDate And B.ProductID = T2.ProductID" & vbCrLf _
         & " Where T1.IsSync = 0 And B.SerialNo is Null And T1.Tag is Null And H.StoreID =" & vStoreID
      
   CN.Execute vSQL
     
   ''''''''''''' Update Remotly Client System '''''''''''''''''''''
   vSQL = "Update " & vConfig & "SaleReturnHeader Set IsSync =1 where IsSync = 0"
   CN.Execute vSQL
   
   ''''''''''''' Delete SaleReturn Body '''''''''''''''''''''
   vSQL = "Delete SaleReturnBody " & vbCrLf _
          & " From SaleReturnBody b" & vbCrLf _
          & " Inner join SaleReturnHeader H on H.SID = B.SID And H.ReturnDate = B.ReturnDate" & vbCrLf _
          & " Left outer join " & vConfig & "SaleReturnHeader T1 on T1.SID = H.Tag And H.StoreID = " & vStoreID & vbCrLf _
          & " Left outer join " & vConfig & "SaleReturnBody T2 on T1.SID = T2.SID And T1.SID = H.Tag And T1.ReturnDate = T2.ReturnDate And T2.ProductID = b.ProductID" & vbCrLf _
          & " WHere H.Tag is Not Null And T2.SerialNo is null And H.StoreID = " & vStoreID
  CN.Execute vSQL
  ''''''''''''' Delete SaleReturn SaleReturnHeader '''''''''''''''''''''
  vSQL = "Delete SaleReturnHeader " & vbCrLf _
          & " From SaleReturnHeader H" & vbCrLf _
          & " Left outer join " & vConfig & "SaleReturnHeader T1 on T1.SID = H.Tag And H.StoreID = " & vStoreID & vbCrLf _
          & " WHere H.Tag is Not Null And T1.SID is null And H.StoreID = " & vStoreID
   CN.Execute vSQL
Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub PurchaseTransaction()
On Error GoTo ErrorHandler
   GetTableColumn ("PurchaseHeader")
   
   ''''''' Make Update Query '''''''
   vSQL = "Update PurchaseHeader Set "
   GridTable.MoveFirst
   For i = 1 To GridTable.Rows
      If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("Tag") Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID") Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
         If UCase(GridTable.Columns("Column_Name").Value) = UCase("StoreID") Then
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = " & vStoreID & ", "
         Else
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
         End If
      End If
      GridTable.MoveNext
   Next i
   GridTable.MoveFirst
   
   vSQL = Replace(vSQL, "''", "Null")
   vSQL = Left(vSQL, Len(vSQL) - 2)
   vSQL = vSQL & vbCrLf _
   & "From PurchaseHeader t1 " & vbCrLf _
   & "Inner Join " & vbCrLf _
   & "(" & vbCrLf _
   & " Select SID, " & vColumnList1 & vbCrLf _
   & " from " & vConfig & "PurchaseHeader " & vbCrLf _
   & " Where IsSync = 0 And Tag is null" & vbCrLf _
   & " )T2"
   vSQL = vSQL & vbCrLf _
   & " on T1.Tag = T2.SID And T1.PurchaseDate = T2.PurchaseDate And T1.StoreID = " & vStoreID & vbCrLf _
   & " Where T1.Tag is not Null "
   CN.Execute vSQL
         
      
   ''''''' Insert record '''''''''''''''
   vColumnList2 = Replace(UCase(vColumnList2), "TAG", "SID")
   vColumnList2 = Replace(UCase(vColumnList2), UCase("T2.StoreID"), vStoreID)
   
   vSQL = "Insert into " & " PurchaseHeader " & "(" & vColumnList1 & ")" & vbCrLf _
         & " Select " & vColumnList2 & vbCrLf _
         & " from " & vConfig & "PurchaseHeader T2" & vbCrLf _
         & " Left outer join  " & " PurchaseHeader T1 "
   vSQL = vSQL & vbCrLf _
         & " on T1.Tag = T2.SID And T1.PurchaseDate = T2.PurchaseDate And T1.StoreID = " & vStoreID & vbCrLf _
         & " Where T1.Tag is Null And T2.IsSync = 0 "
    CN.Execute vSQL
   
           
   GetTableColumn ("PurchaseBody")
   
    ''''''' Make Update Query '''''''''''''''''''''''
   vSQL = "Update PurchaseBody Set "
   
   GridTable.MoveFirst
   For i = 1 To GridTable.Rows
      If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("SID") Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID") Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
         If UCase(GridTable.Columns("Column_Name").Value) = UCase("StoreID") Then
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = " & vStoreID & ", "
         Else
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
         End If
      End If
      GridTable.MoveNext
   Next i
   GridTable.MoveFirst
   
   vSQL = Replace(vSQL, "''", "Null")
   vSQL = Left(vSQL, Len(vSQL) - 2)
   vSQL = vSQL & vbCrLf _
   & "From PurchaseBody B " & vbCrLf _
   & "inner join PurchaseHeader H on H.SID = B.SID" & vbCrLf _
   & "Inner Join " & vbCrLf _
   & "(" & vbCrLf _
   & " Select " & vColumnList2 & vbCrLf _
   & " from " & vConfig & "PurchaseBody T2 " & vbCrLf _
   & " inner join " & vConfig & "PurchaseHeader T1 " & vbCrLf _
   & " on T1.SID = T2.SID And T1.PurchaseDate = T2.PurchaseDate " & vbCrLf _
   & " inner join  PurchaseHeader H on H.Tag = T1.SID " & vbCrLf _
   & " Where T1.Tag is Null And T1.IsSync = 0 And H.StoreID = " & vStoreID & vbCrLf _
   & " )T2"
   vSQL = vSQL & vbCrLf _
   & " on H.Tag = T2.SID And H.PurchaseDate = T2.PurchaseDate And T2.ProductID = B.ProductID And T1.StoreID = " & vStoreID & vbCrLf _
   & " Where H.Tag is not Null "
   CN.Execute vSQL

   ''''''' Insert record '''''''''''
          
   vColumnList2 = Replace(UCase(vColumnList2), "T2.SID", "H.SID")
   vColumnList2 = Replace(UCase(vColumnList2), UCase("T2.StoreID"), vStoreID)
   vSQL = "Insert into " & " PurchaseBody " & "(" & vColumnList1 & ")" & vbCrLf _
         & " Select " & vColumnList2 & vbCrLf _
         & " from " & vConfig & "PurchaseBody T2" & vbCrLf _
         & " inner join " & vConfig & "PurchaseHeader T1" & vbCrLf _
         & " on T2.SID = T1.SID And T2.PurchaseDate = T1.PurchaseDate" & vbCrLf _
         & " inner join  PurchaseHeader H on H.Tag = T1.SID And H.StoreID =" & vStoreID & vbCrLf _
         & " left outer Join PurchaseBody B on B.SID = H.SID and B.PurID = T2.PurID and B.PurchaseDate = T2.PurchaseDate And B.ProductID = T2.ProductID" & vbCrLf _
         & " Where T1.IsSync = 0 And B.SerialNo is Null And T1.Tag is Null And H.StoreID =" & vStoreID
      
   CN.Execute vSQL
     
   ''''''''''''' Update Remotly Client System '''''''''''''''''''''
   vSQL = "Update " & vConfig & "PurchaseHeader Set IsSync =1 where IsSync = 0"
   CN.Execute vSQL
   
   ''''''''''''' Delete Purchase Body '''''''''''''''''''''
   vSQL = "Delete PurchaseBody " & vbCrLf _
          & " From PurchaseBody b" & vbCrLf _
          & " Inner join PurchaseHeader H on H.SID = B.SID And H.PurchaseDate = B.PurchaseDate" & vbCrLf _
          & " Left outer join " & vConfig & "PurchaseHeader T1 on T1.SID = H.Tag And H.StoreID = " & vStoreID & vbCrLf _
          & " Left outer join " & vConfig & "PurchaseBody T2 on T1.SID = T2.SID And T1.SID = H.Tag And T1.PurchaseDate = T2.PurchaseDate And T2.ProductID = b.ProductID" & vbCrLf _
          & " WHere H.Tag is Not Null And T2.SerialNo is null "
  CN.Execute vSQL
  ''''''''''''' Delete Purchase PurchaseHeader '''''''''''''''''''''
  vSQL = "Delete PurchaseHeader " & vbCrLf _
          & " From PurchaseHeader H" & vbCrLf _
          & " Left outer join " & vConfig & "PurchaseHeader T1 on T1.SID = H.Tag And H.StoreID = " & vStoreID & vbCrLf _
          & " WHere H.Tag is Not Null And T1.SID is null And H.StoreID = " & vStoreID
   CN.Execute vSQL
Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
   
End Sub

Private Sub StockTransferExport()
   On Error GoTo ErrorHandler
   Dim i As Integer, j As Integer
   Dim vMainStoreID As Byte
  
   '''''''''''''''''''' Purchase ''''''''''''''''''''
   vSQL = "Select * From StockTransferHeader " & vbCrLf _
         & "where isSync = 0 and Tag is null and FromStoreID = 1 and ToStoreID = " & vStoreID
         With CN.Execute(vSQL)
            While Not .EOF
                  
               FunGetMaxID = CN.Execute("Select isnull(max(PurID),0)+1 from " & vConfig & "PurchaseHeader").Fields(0).Value
               
               ''''''' Insert record
'               vSQL = "Insert into " & vConfig & "PurchaseHeader (PurID, PurchaseDate, VendorID, Tag, BillNo, TotalAmount, OtherCharges, UserNo, StoreID ) " & vbCrLf _
                      & " Values " & vbCrLf _
                      & "(" & FunGetMaxID & ",'" & .Fields("TransferDate").Value & "'," & 611 & "," & .Fields("TransferID").Value & "," & IIf(IsNull(.Fields("BillNo").Value), 0, .Fields("BillNO").Value) & "," & IIf(IsNull(.Fields("TotalAmount").Value), 0, .Fields("TotalAmount").Value) & "," & IIf(IsNull(.Fields("OtherChargesVal").Value), 0, .Fields("OtherChargesVal").Value) & "," & .Fields("UserNo").Value & ",1)"
               
               vSQL = "Exec " & vConfig & "ProdConvertStockIntoPurchase " & FunGetMaxID & ",'" & .Fields("TransferDate").Value & "'," & 611 & "," & .Fields("TransferID").Value & "," & IIf(IsNull(.Fields("BillNo").Value), 0, .Fields("BillNO").Value) & "," & IIf(IsNull(.Fields("TotalAmount").Value), 0, .Fields("TotalAmount").Value) & "," & IIf(IsNull(.Fields("OtherChargesVal").Value), 0, .Fields("OtherChargesVal").Value) & "," & .Fields("UserNo").Value & ",1"
               vSID = CN.Execute(vSQL).Fields(0).Value
               
               ''''''''''' Insert Purchase Body
               vSQL = "Select * from StockTransferBody " & vbCrLf _
               & "Where TransferID = " & .Fields("TransferID").Value & " and TransferDate = '" & .Fields("TransferDate").Value & "'"
               If Rs.State = adStateOpen Then Rs.Close
               Rs.Open vSQL, CN, adOpenStatic, adLockBatchOptimistic
                  While Not Rs.EOF
                     vSQL = "Insert into " & vConfig & "PurchaseBody (SID, PurID, PurchaseDate,  Code, ProductID, QtyPack, QtyLoose, Multiplier, Price, DiscVal, DiscPer, DiscPC, Amount ) Values " & vbCrLf _
                           & "(" & vSID & "," & FunGetMaxID & ",'" & Rs!TransferDate & "'," & Rs!Code & ",'" & Rs!ProductID & "'," & IIf(IsNull(Rs!QtyPack), "Null", Rs!QtyPack) & "," & Rs!QtyLoose & "," & IIf(IsNull(Rs!Multiplier), "Null", Rs!Multiplier) & "," & IIf(IsNull(Rs!Price), 0, Rs!Price) & ",0,0,0," & IIf(IsNull(Rs!Amount), 0, Rs!Amount) & ")"
                     CN.Execute vSQL
                     Rs.MoveNext
                  Wend
                     CN.Execute ("Update StockTransferHeader Set Tag = " & vSID & ", BranchID =" & vSID & ", BranchDate = TransferDate, IsSync = 1 Where IsSync = 0 and tag is null and TransferID = " & .Fields("TransferID").Value & " and TransferDate = '" & .Fields("TransferDate").Value & "' and  ToStoreID = " & vStoreID)
               .MoveNext
            Wend
         End With
         
         vSQL = "Select * From StockTransferHeader " & vbCrLf _
         & "where isSync = 0 and FromStoreID = 1 and ToStoreID = " & vStoreID
         With CN.Execute(vSQL)
            If .EOF = False Then
               '''''''''''''''''''''' Update PurchaseBody
                vSQL = "Update " & vConfig & "PurchaseBody Set " & vbCrLf _
                & "Multiplier = STB.Multiplier, QtyPack = STB.QtyPack, QtyLoose = STB.QtyLoose, DiscPC = 0, DiscPer=0, DiscVal=0, Price = STB.Price, Amount = STB.Amount " & vbCrLf _
                & "From " & vConfig & "PurchaseBody PB" & vbCrLf _
                & "inner join  StockTransferHeader STH on STH.BranchID = PB.SID and STH.BranchDate = PB.PurchaseDate And STH.ToStoreID = " & vStoreID & vbCrLf _
                & "Inner jOIN StockTransferbody STB on STB.TransferID = STH.TransferID and STB.TransferDate  = STH.TransferDate and STB.ProductID = PB.productID " & vbCrLf _
                & "Where IsSync = 0 and STH.FromStoreID = 1 And STH.ToStoreID = " & vStoreID
                
                CN.Execute vSQL
                
               ''''''''''' Insert Purchase Body
               vSQL = "Select STB.*, STH.Tag, PH.SID, PH.PurID from StockTransferBody STB Inner Join StockTransferHeader STH On STH.TransferID = STB.TransferID And STH.TransferDate = STB.TransferDate " & vbCrLf _
                     & " Inner Join " & vConfig & "PurchaseHeader PH on PH.SID = STH.SID And PH.PurchaseDate = STH.BranchDate" & vbCrLf _
                     & " Left Outer Join " & vConfig & "PurchaseBody PB on PB.SID = PH.SID And PB.PurchaseDate = PH.PurchaseDate And PB.ProductID = STB.ProductID " & vbCrLf _
                     & " Where STH.Tag Is Not Null And STH.IsSync = 0 And PB.SerialNo is Null and STH.FromStoreID = 1 And STH.ToStoreID =" & vStoreID
               If Rs.State = adStateOpen Then Rs.Close
               Rs.Open vSQL, CN, adOpenStatic, adLockBatchOptimistic
                  While Not Rs.EOF
                     vSQL = "Insert into " & vConfig & "PurchaseBody (SID, PurID, PurchaseDate,  Code, ProductID, QtyPack, QtyLoose, Multiplier, Price, DiscVal, DiscPer, DiscPC, Amount ) Values " & vbCrLf _
                           & "(" & Rs!SID & "," & Rs!PurID & ",'" & Rs!TransferDate & "'," & Rs!Code & ",'" & Rs!ProductID & "'," & Rs!QtyPack & "," & Rs!QtyLoose & "," & Rs!Multiplier & "," & IIf(IsNull(Rs!Price), 0, Rs!Price) & ",0,0,0," & IIf(IsNull(Rs!Amount), 0, Rs!Amount) & ")"
                     CN.Execute vSQL
                     Rs.MoveNext
                  Wend
                
                '''''''''''''''''''''' Update PurchaseHeader
                vSQL = "Update " & vConfig & "PurchaseHeader Set " & vbCrLf _
                & "TotalAmount = STH.TotalAmount" & vbCrLf _
                & "From " & vConfig & "PurchaseHeader PH" & vbCrLf _
                & "inner join  StockTransferHeader STH on STH.BranchID = PH.SID and STH.BranchDate = PH.PurchaseDate And STH.ToStoreID = " & vStoreID & vbCrLf _
                & "Where IsSync = 0 and STH.FromStoreID = 1 And STH.ToStoreID = " & vStoreID
                CN.Execute vSQL
                
                ''''''''''''' Update StockTransferHeader '''''''''''''''''''''
                CN.Execute ("Update StockTransferHeader Set IsSync = 1 Where IsSync = 0 And FromStoreID = 1  and  ToStoreID = " & vStoreID)
            End If
         End With
         
         ''''''''''''' Delete Purchase Body '''''''''''''''''''''
                vSQL = "Delete " & vConfig & "PurchaseBody" & vbCrLf _
                       & " From " & vConfig & "PurchaseBody PB" & vbCrLf _
                       & " Inner join " & vConfig & "PurchaseHeader PH on PH.SID = PB.SID and PH.PurchaseDate = PB.PurchaseDate" & vbCrLf _
                       & " Left outer join StockTransferHeader STH on STH.BranchID = PH.SID and STH.BranchDate = PH.PurchaseDate And STH.ToStoreID = " & vStoreID & vbCrLf _
                       & " Left outer join StockTransferbody STB on STB.TransferID = STH.TransferID And STB.TransferDate = STH.TransferDate And STB.ProductID = PB.ProductID" & vbCrLf _
                       & " WHere PH.Tag is Not null And STB.SerialNo is null And PH.Tag Not like '%Date%'"
               CN.Execute vSQL
         
         ''''''''''''' Delete Sale PurchaseHeader '''''''''''''''''''''
         vSQL = "Delete " & vConfig & "PurchaseHeader " & vbCrLf _
                & "From " & vConfig & "PurchaseHeader PH" & vbCrLf _
                & "Left Outer join  StockTransferHeader STH on STH.BranchID = PH.SID and STH.BranchDate = PH.PurchaseDate And STH.ToStoreID = " & vStoreID & vbCrLf _
                & "Where PH.Tag is Not null and STH.TransferID is null And PH.Tag Not like '%Date%'"
         CN.Execute vSQL
         
      '''''''''''''''  ''''''''''''''''  '''''''''''''''  ''''''''''''''''  '''''''''''''''  ''''''''''''''''
      '''''''''''''''  ''''''''''''''''  '''''''''''''''  ''''''''''''''''  '''''''''''''''  ''''''''''''''''
      '''''''''''''''  ''''''''''''''''  '''''''''''''''  ''''''''''''''''  '''''''''''''''  ''''''''''''''''
      
      '''''''''''''''''''' Purchase Return ''''''''''''''''''''
   vSQL = "Select * From StockTransferHeader " & vbCrLf _
         & "where isSync = 0 and Tag is null and ToStoreID = 1 and FromStoreID = " & vStoreID
         With CN.Execute(vSQL)
            While Not .EOF
                  
               FunGetMaxID = CN.Execute("Select isnull(max(ReturnID),0)+1 from " & vConfig & "PurchaseReturnHeader").Fields(0).Value
               
               ''''''' Insert record
'               vSQL = "Insert into " & vConfig & "PurchaseReturnHeader (ReturnID, ReturnDate, VendorID, Tag, BillNo, TotalAmount, OtherCharges, UserNo, StoreID ) Values " & vbCrLf _
'                      & "(" & FunGetMaxID & ",'" & .Fields("TransferDate").Value & "'," & 611 & "," & .Fields("TransferID").Value & "," & IIf(IsNull(.Fields("BillNo").Value), 0, .Fields("BillNO").Value) & "," & IIf(IsNull(.Fields("TotalAmount").Value), 0, .Fields("TotalAmount").Value) & "," & IIf(IsNull(.Fields("OtherChargesVal").Value), 0, .Fields("OtherChargesVal").Value) & "," & .Fields("UserNo").Value & ",1)"
'               CN.Execute vSQL
               
               vSQL = "Exec " & vConfig & "ProdConvertStockIntoPurchaseReturn " & FunGetMaxID & ",'" & .Fields("TransferDate").Value & "'," & 611 & "," & .Fields("TransferID").Value & "," & IIf(IsNull(.Fields("BillNo").Value), 0, .Fields("BillNO").Value) & "," & IIf(IsNull(.Fields("TotalAmount").Value), 0, .Fields("TotalAmount").Value) & "," & IIf(IsNull(.Fields("OtherChargesVal").Value), 0, .Fields("OtherChargesVal").Value) & "," & .Fields("UserNo").Value & ",1"
               vSID = CN.Execute(vSQL).Fields(0).Value
               
               ''''''''''' Insert PurchaseReturn Body
               vSQL = "Select * from StockTransferBody " & vbCrLf _
               & "Where TransferID = " & .Fields("TransferID").Value & " and TransferDate = '" & .Fields("TransferDate").Value & "'"
               If Rs.State = adStateOpen Then Rs.Close
               Rs.Open vSQL, CN, adOpenStatic, adLockBatchOptimistic
                  While Not Rs.EOF
                     vSQL = "Insert into " & vConfig & "PurchaseReturnBody (SID, ReturnID, ReturnDate,  Code, ProductID, QtyPack, QtyLoose, Multiplier, Price, DiscVal, DiscPer, DiscPC, Amount ) Values " & vbCrLf _
                           & "(" & vSID & "," & FunGetMaxID & ",'" & Rs!TransferDate & "'," & Rs!Code & ",'" & Rs!ProductID & "'," & Rs!QtyPack & "," & Rs!QtyLoose & "," & Rs!Multiplier & "," & IIf(IsNull(Rs!Price), 0, Rs!Price) & ",0,0,0," & IIf(IsNull(Rs!Amount), 0, Rs!Amount) & ")"
                     CN.Execute vSQL
                     Rs.MoveNext
                  Wend
                  vSQL = "Update StockTransferHeader Set Tag = " & vSID & ", BranchID =" & vSID & ", BranchDate = TransferDate, IsSync = 1 Where IsSync = 0 and tag is null and TransferID = " & .Fields("TransferID").Value & " and TransferDate = '" & .Fields("TransferDate").Value & "' and ToStoreID = 1 and FromStoreID = " & vStoreID
                  CN.Execute vSQL
               .MoveNext
            Wend
         End With
         
         vSQL = "Select * From StockTransferHeader " & vbCrLf _
         & "where isSync = 0 and ToStoreID = 1 and FromStoreID = " & vStoreID
         With CN.Execute(vSQL)
            If .EOF = False Then
               '''''''''''''''''''''' Update PurchaseReturnBody
                vSQL = "Update " & vConfig & "PurchaseReturnBody Set " & vbCrLf _
                & "Multiplier = STB.Multiplier, QtyPack = STB.QtyPack, QtyLoose = STB.QtyLoose, DiscPC = 0, DiscPer=0, DiscVal=0, Price = STB.Price, Amount = STB.Amount " & vbCrLf _
                & "From " & vConfig & "PurchaseReturnBody PRB" & vbCrLf _
                & "inner join StockTransferHeader STH on STH.BranchID = PRB.SID and STH.BranchDate = PRB.ReturnDate And STH.FromStoreID = " & vStoreID & vbCrLf _
                & "Inner join StockTransferbody STB on STB.TransferID = STH.TransferID and STB.TransferDate = STH.TransferDate and STB.ProductID = PRB.productID " & vbCrLf _
                & "Where IsSync = 0 and STH.ToStoreID = 1 And STH.FromStoreID = " & vStoreID
                
                CN.Execute vSQL
                
               ''''''''''' Insert PurchaseReturn Body
               vSQL = "Select STB.*, STH.Tag, PRH.SID, PRH.ReturnID from StockTransferBody STB Inner Join StockTransferHeader STH On STH.TransferID = STB.TransferID And STH.TransferDate = STB.TransferDate " & vbCrLf _
                     & " Inner Join " & vConfig & "PurchaseReturnHeader PRH on PRH.ReturnID = STH.BranchID And PRH.ReturnDate = STH.BranchDate" & vbCrLf _
                     & " Left Outer Join " & vConfig & "PurchaseReturnBody PRB on PRB.ReturnID = PRH.ReturnID And PRB.ReturnDate = PRH.ReturnDate And STH.Tag = PRH.ReturnID And PRB.ProductID = STB.ProductID " & vbCrLf _
                     & " Where STH.Tag Is Not Null And STH.IsSync = 0 And PRB.SerialNo is Null and STH.ToStoreID = 1 And STH.FromStoreID =" & vStoreID
               If Rs.State = adStateOpen Then Rs.Close
               Rs.Open vSQL, CN, adOpenStatic, adLockBatchOptimistic
                  While Not Rs.EOF
                     vSQL = "Insert into " & vConfig & "PurchaseReturnBody (SID, ReturnID, ReturnDate,  Code, ProductID, QtyPack, QtyLoose, Multiplier, Price, DiscVal, DiscPer, DiscPC, Amount ) Values " & vbCrLf _
                           & "(" & Rs!SID & "," & Rs!ReturnID & ",'" & Rs!TransferDate & "'," & Rs!Code & ",'" & Rs!ProductID & "'," & Rs!QtyPack & "," & Rs!QtyLoose & "," & Rs!Multiplier & "," & IIf(IsNull(Rs!Price), 0, Rs!Price) & ",0,0,0," & IIf(IsNull(Rs!Amount), 0, Rs!Amount) & ")"
                     CN.Execute vSQL
                     Rs.MoveNext
                  Wend
                
                '''''''''''''''''''''' Update PurchaseReturnHeader
                vSQL = "Update " & vConfig & "PurchaseReturnHeader Set " & vbCrLf _
                & "TotalAmount = STH.TotalAmount" & vbCrLf _
                & "From " & vConfig & "PurchaseReturnHeader PRH" & vbCrLf _
                & "inner join  StockTransferHeader STH on STH.BranchID = PRH.SID and STH.BranchDate = PRH.ReturnDate And STH.FromStoreID = " & vStoreID & vbCrLf _
                & "Where IsSync = 0 and STH.TOStoreID = 1 And STH.FromStoreID = " & vStoreID
                CN.Execute vSQL
                
                ''''''''''''' Update StockTransferHeader '''''''''''''''''''''
                CN.Execute ("Update StockTransferHeader Set IsSync = 1 Where IsSync = 0 And ToStoreID = 1  and  FromStoreID = " & vStoreID)
            End If
         End With
         
         ''''''''''''' Delete PurchaseReturn Body '''''''''''''''''''''
          vSQL = "Delete " & vConfig & "PurchaseReturnBody" & vbCrLf _
                 & " From " & vConfig & "PurchaseReturnBody PRB" & vbCrLf _
                 & " Inner join " & vConfig & "PurchaseReturnHeader PRH on PRH.SID = PRB.SID and PRH.ReturnDate = PRB.ReturnDate" & vbCrLf _
                 & " Left outer join StockTransferHeader STH on STH.BranchID = PRH.SID and STH.BranchDate = PRH.ReturnDate And STH.FromStoreID = " & vStoreID & vbCrLf _
                 & " Left outer join StockTransferbody STB on STB.TransferID = STH.TransferID And STB.TransferDate = STH.TransferDate And STB.ProductID = PRB.ProductID" & vbCrLf _
                 & " Where PRH.Tag is Not null And STB.SerialNo is null And PRH.Tag Not like '%Date%'"
         CN.Execute vSQL
         
         ''''''''''''' Delete Sale PurchaseReturnHeader '''''''''''''''''''''
         vSQL = "Delete " & vConfig & "PurchaseReturnHeader " & vbCrLf _
                & "From " & vConfig & "PurchaseReturnHeader PRH" & vbCrLf _
                & "Left Outer join  StockTransferHeader STH on STH.BranchID = PRH.SID and STH.BranchDate = PRH.ReturnDate And STH.FromStoreID = " & vStoreID & vbCrLf _
                & "Where PRH.Tag is Not null and STH.TransferID is null and STH.BranchID = PRH.ReturnID And PRH.Tag Not like '%Date%'"
         CN.Execute vSQL
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CreditVouchersTransaction()
On Error GoTo ErrorHandler
   GetTableColumn ("CreditVouchers")
   
   ''''''' Make Update Query
   vSQL = "Update CreditVouchers Set "
   GridTable.MoveFirst
   For i = 1 To GridTable.Rows
      If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("Tag") Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID") Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
         If UCase(GridTable.Columns("Column_Name").Value) = UCase("StoreID") Then
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = " & vStoreID & ", "
         Else
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
         End If
      End If
      GridTable.MoveNext
   Next i
   GridTable.MoveFirst
   
   vSQL = Replace(vSQL, "''", "Null")
   vSQL = Left(vSQL, Len(vSQL) - 2)
   vSQL = vSQL & vbCrLf _
   & "From CreditVouchers t1 " & vbCrLf _
   & "Inner Join " & vbCrLf _
   & "(" & vbCrLf _
   & " Select SID, " & vColumnList1 & vbCrLf _
   & " from " & vConfig & "CreditVouchers " & vbCrLf _
   & " Where IsSync = 0 And Tag is null" & vbCrLf _
   & " )T2"
   vSQL = vSQL & vbCrLf _
   & " on T1.Tag = T2.SID And T1.VoucherDate = T2.VoucherDate And T1.StoreID = " & vStoreID & vbCrLf _
   & " Where T1.Tag is not Null "
   CN.Execute vSQL
         
   
   
   ''''''' Insert record
   vColumnList2 = Replace(UCase(vColumnList2), "TAG", "SID")
   vColumnList2 = Replace(UCase(vColumnList2), UCase("T2.StoreID"), vStoreID)
   
   vSQL = "Insert into " & " CreditVouchers " & "(" & vColumnList1 & ")" & vbCrLf _
         & " Select " & vColumnList2 & vbCrLf _
         & " from " & vConfig & "CreditVouchers T2" & vbCrLf _
         & " Left outer join  " & " CreditVouchers T1 "
   vSQL = vSQL & vbCrLf _
         & " on T1.Tag = T2.SID And T1.VoucherDate = T2.VoucherDate And T1.StoreID = " & vStoreID & vbCrLf _
         & " Where T1.Tag is Null And T2.IsSync = 0 "
    CN.Execute vSQL
   
  
         
   GetTableColumn ("CreditVouchersBody")
   
    ''''''' Make Update Query
   vSQL = "Update CreditVouchersBody Set "
   
   GridTable.MoveFirst
   For i = 1 To GridTable.Rows
      If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("SID") Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID") Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
         If UCase(GridTable.Columns("Column_Name").Value) = UCase("StoreID") Then
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = " & vStoreID & ", "
         Else
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
         End If
      End If
      GridTable.MoveNext
   Next i
   GridTable.MoveFirst
   
   vSQL = Replace(vSQL, "''", "Null")
   vSQL = Left(vSQL, Len(vSQL) - 2)
   vSQL = vSQL & vbCrLf _
   & "From CreditVouchersBody B " & vbCrLf _
   & "inner join CreditVouchers H on H.SID = B.SID" & vbCrLf _
   & "Inner Join " & vbCrLf _
   & "(" & vbCrLf _
   & " Select " & vColumnList2 & vbCrLf _
   & " from " & vConfig & "CreditVouchersBody T2 " & vbCrLf _
   & " inner join " & vConfig & "CreditVouchers T1 " & vbCrLf _
   & " on T1.SID = T2.SID " & vbCrLf _
   & " inner join  CreditVouchers H on H.Tag = T1.SID And H.StoreID = " & vStoreID & vbCrLf _
   & " Where T1.Tag is Null And T1.IsSync = 0 And H.StoreID = " & vStoreID & vbCrLf _
   & " )T2"
   vSQL = vSQL & vbCrLf _
   & " on H.Tag = T2.SID And T2.AccountNo = B.AccountNo And H.StoreID = " & vStoreID & vbCrLf _
   & " Where H.Tag is not Null "
   CN.Execute vSQL

   ''''''' Insert record
          
   vColumnList2 = Replace(UCase(vColumnList2), "T2.SID", "H.SID")
   vColumnList2 = Replace(UCase(vColumnList2), UCase("T2.StoreID"), vStoreID)
   
   vSQL = "Insert into " & " CreditVouchersBody " & "(" & vColumnList1 & ")" & vbCrLf _
         & " Select " & vColumnList2 & vbCrLf _
         & " from " & vConfig & "CreditVouchersBody T2" & vbCrLf _
         & " inner join " & vConfig & "CreditVouchers T1" & vbCrLf _
         & " on T2.SID = T1.SID " & vbCrLf _
         & " inner join  CreditVouchers H on H.Tag = T1.SID And H.StoreID =" & vStoreID & vbCrLf _
         & " left outer Join CreditVouchersBody B on B.SID = H.SID And B.AccountNo = T2.AccountNo" & vbCrLf _
         & " Where T1.IsSync = 0 And B.SerialNo is Null And (T1.Tag is Null or T1.Tag = '') And H.StoreID =" & vStoreID
      
   CN.Execute vSQL
     
   ''''''''''''' Update Remotly Client System '''''''''''''''''''''
   vSQL = "Update " & vConfig & "CreditVouchers Set IsSync =1 where IsSync = 0"
   CN.Execute vSQL
   
   ''''''''''''' Delete Credit Body '''''''''''''''''''''
   vSQL = "Delete CreditVouchersBody " & vbCrLf _
          & " From CreditVouchersBody b" & vbCrLf _
          & " Inner join CreditVouchers H on H.SID = B.SID " & vbCrLf _
          & " Left outer join " & vConfig & "CreditVouchers T1 on T1.SID = H.Tag And H.StoreID = " & vStoreID & vbCrLf _
          & " Left outer join " & vConfig & "CreditVouchersBody T2 on T1.SID = T2.SID And T1.SID = H.Tag And T2.AccountNo = b.AccountNo" & vbCrLf _
          & " WHere H.Tag is Not Null And T2.SerialNo is null And H.StoreID = " & vStoreID
  CN.Execute vSQL
  ''''''''''''' Delete Credit CreditVouchers '''''''''''''''''''''
  vSQL = "Delete CreditVouchers " & vbCrLf _
          & " From CreditVouchers H" & vbCrLf _
          & " Left outer join " & vConfig & "CreditVouchers T1 on T1.SID = H.Tag And H.StoreID = " & vStoreID & vbCrLf _
          & " WHere H.Tag is Not Null And T1.SID is null And H.StoreID = " & vStoreID
   CN.Execute vSQL
Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
   
End Sub

Private Sub DebitVouchersTransaction()
On Error GoTo ErrorHandler
   GetTableColumn ("DebitVouchers")
   
   ''''''' Make Update Query
   vSQL = "Update DebitVouchers Set "
   GridTable.MoveFirst
   For i = 1 To GridTable.Rows
      If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("Tag") Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID") Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
         If UCase(GridTable.Columns("Column_Name").Value) = UCase("StoreID") Then
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = " & vStoreID & ", "
         Else
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
         End If
      End If
      GridTable.MoveNext
   Next i
   GridTable.MoveFirst
   
   vSQL = Replace(vSQL, "''", "Null")
   vSQL = Left(vSQL, Len(vSQL) - 2)
   vSQL = vSQL & vbCrLf _
   & "From DebitVouchers t1 " & vbCrLf _
   & "Inner Join " & vbCrLf _
   & "(" & vbCrLf _
   & " Select SID, " & vColumnList1 & vbCrLf _
   & " from " & vConfig & "DebitVouchers " & vbCrLf _
   & " Where IsSync = 0 And Tag is null" & vbCrLf _
   & " )T2"
   vSQL = vSQL & vbCrLf _
   & " on T1.Tag = T2.SID And T1.VoucherDate = T2.VoucherDate And T1.StoreID = " & vStoreID & vbCrLf _
   & " Where T1.Tag is not Null "
   CN.Execute vSQL
         
   
   
   ''''''' Insert record
   vColumnList2 = Replace(UCase(vColumnList2), "TAG", "SID")
   vColumnList2 = Replace(UCase(vColumnList2), UCase("T2.StoreID"), vStoreID)
   
   vSQL = "Insert into " & " DebitVouchers " & "(" & vColumnList1 & ")" & vbCrLf _
         & " Select " & vColumnList2 & vbCrLf _
         & " from " & vConfig & "DebitVouchers T2" & vbCrLf _
         & " Left outer join  " & " DebitVouchers T1 "
   vSQL = vSQL & vbCrLf _
         & " on T1.Tag = T2.SID And T1.VoucherDate = T2.VoucherDate And T1.StoreID = " & vStoreID & vbCrLf _
         & " Where T1.Tag is Null And T2.IsSync = 0 "
    CN.Execute vSQL
   
  
         
   GetTableColumn ("DebitVouchersBody")
   
    ''''''' Make Update Query
   vSQL = "Update DebitVouchersBody Set "
   
   GridTable.MoveFirst
   For i = 1 To GridTable.Rows
      If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("SID") Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID") Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
         If UCase(GridTable.Columns("Column_Name").Value) = UCase("StoreID") Then
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = " & vStoreID & ", "
         Else
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
         End If
      End If
      GridTable.MoveNext
   Next i
   GridTable.MoveFirst
   
   vSQL = Replace(vSQL, "''", "Null")
   vSQL = Left(vSQL, Len(vSQL) - 2)
   vSQL = vSQL & vbCrLf _
   & "From DebitVouchersBody B " & vbCrLf _
   & "inner join DebitVouchers H on H.SID = B.SID" & vbCrLf _
   & "Inner Join " & vbCrLf _
   & "(" & vbCrLf _
   & " Select " & vColumnList2 & vbCrLf _
   & " from " & vConfig & "DebitVouchersBody T2 " & vbCrLf _
   & " inner join " & vConfig & "DebitVouchers T1 " & vbCrLf _
   & " on T1.SID = T2.SID " & vbCrLf _
   & " inner join  DebitVouchers H on H.Tag = T1.SID And H.StoreID = " & vStoreID & vbCrLf _
   & " Where T1.Tag is Null And T1.IsSync = 0 And H.StoreID = " & vStoreID & vbCrLf _
   & " )T2"
   vSQL = vSQL & vbCrLf _
   & " on H.Tag = T2.SID And T2.AccountNo = B.AccountNo And H.StoreID = " & vStoreID & vbCrLf _
   & " Where H.Tag is not Null "
   CN.Execute vSQL

   ''''''' Insert record
          
   vColumnList2 = Replace(UCase(vColumnList2), "T2.SID", "H.SID")
   vColumnList2 = Replace(UCase(vColumnList2), UCase("T2.StoreID"), vStoreID)
   
   vSQL = "Insert into " & " DebitVouchersBody " & "(" & vColumnList1 & ")" & vbCrLf _
         & " Select " & vColumnList2 & vbCrLf _
         & " from " & vConfig & "DebitVouchersBody T2" & vbCrLf _
         & " inner join " & vConfig & "DebitVouchers T1" & vbCrLf _
         & " on T2.SID = T1.SID " & vbCrLf _
         & " inner join  DebitVouchers H on H.Tag = T1.SID And H.StoreID =" & vStoreID & vbCrLf _
         & " left outer Join DebitVouchersBody B on B.SID = H.SID And B.AccountNo = T2.AccountNo" & vbCrLf _
         & " Where T1.IsSync = 0 And B.SerialNo is Null And (T1.Tag is Null or T1.Tag = '') And H.StoreID =" & vStoreID
      
   CN.Execute vSQL
     
   ''''''''''''' Update Remotly Client System '''''''''''''''''''''
   vSQL = "Update " & vConfig & "DebitVouchers Set IsSync =1 where IsSync = 0"
   CN.Execute vSQL
   
   ''''''''''''' Delete Debit Body '''''''''''''''''''''
   vSQL = "Delete DebitVouchersBody " & vbCrLf _
          & " From DebitVouchersBody b" & vbCrLf _
          & " Inner join DebitVouchers H on H.SID = B.SID " & vbCrLf _
          & " Left outer join " & vConfig & "DebitVouchers T1 on T1.SID = H.Tag And H.StoreID = " & vStoreID & vbCrLf _
          & " Left outer join " & vConfig & "DebitVouchersBody T2 on T1.SID = T2.SID And T1.SID = H.Tag And T2.AccountNo = b.AccountNo" & vbCrLf _
          & " WHere H.Tag is Not Null And T2.SerialNo is null And H.StoreID = " & vStoreID
  CN.Execute vSQL
  ''''''''''''' Delete Debit DebitVouchers '''''''''''''''''''''
  vSQL = "Delete DebitVouchers " & vbCrLf _
          & " From DebitVouchers H" & vbCrLf _
          & " Left outer join " & vConfig & "DebitVouchers T1 on T1.SID = H.Tag And H.StoreID = " & vStoreID & vbCrLf _
          & " WHere H.Tag is Not Null And T1.SID is null And H.StoreID = " & vStoreID
   CN.Execute vSQL
Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
   
End Sub


Private Sub JournalVouchersTransaction()
On Error GoTo ErrorHandler
   GetTableColumn ("JournalVouchers")
   
   ''''''' Make Update Query
   vSQL = "Update JournalVouchers Set "
   GridTable.MoveFirst
   For i = 1 To GridTable.Rows
      If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("Tag") Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID") Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
         If UCase(GridTable.Columns("Column_Name").Value) = UCase("StoreID") Then
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = " & vStoreID & ", "
         Else
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
         End If
      End If
      GridTable.MoveNext
   Next i
   GridTable.MoveFirst
   
   vSQL = Replace(vSQL, "''", "Null")
   vSQL = Left(vSQL, Len(vSQL) - 2)
   vSQL = vSQL & vbCrLf _
   & "From JournalVouchers t1 " & vbCrLf _
   & "Inner Join " & vbCrLf _
   & "(" & vbCrLf _
   & " Select SID, " & vColumnList1 & vbCrLf _
   & " from " & vConfig & "JournalVouchers " & vbCrLf _
   & " Where IsSync = 0 And Tag is null" & vbCrLf _
   & " )T2"
   vSQL = vSQL & vbCrLf _
   & " on T1.Tag = T2.SID And T1.VoucherDate = T2.VoucherDate And T1.StoreID = " & vStoreID & vbCrLf _
   & " Where T1.Tag is not Null "
   CN.Execute vSQL
         
   
   
   ''''''' Insert record
   vColumnList2 = Replace(UCase(vColumnList2), "TAG", "SID")
   vColumnList2 = Replace(UCase(vColumnList2), UCase("T2.StoreID"), vStoreID)
   
   vSQL = "Insert into " & " JournalVouchers " & "(" & vColumnList1 & ")" & vbCrLf _
         & " Select " & vColumnList2 & vbCrLf _
         & " from " & vConfig & "JournalVouchers T2" & vbCrLf _
         & " Left outer join  " & " JournalVouchers T1 "
   vSQL = vSQL & vbCrLf _
         & " on T1.Tag = T2.SID And T1.VoucherDate = T2.VoucherDate And T1.StoreID = " & vStoreID & vbCrLf _
         & " Where T1.Tag is Null And T2.IsSync = 0 "
    CN.Execute vSQL
   
  
         
   GetTableColumn ("JournalVouchersBody")
   
    ''''''' Make Update Query
   vSQL = "Update JournalVouchersBody Set "
   
   GridTable.MoveFirst
   For i = 1 To GridTable.Rows
      If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or UCase(GridTable.Columns("Column_Name").Value) = UCase("SID") Or UCase(GridTable.Columns("Column_Name").Value) = UCase("StampID") Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
         If UCase(GridTable.Columns("Column_Name").Value) = UCase("StoreID") Then
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = " & vStoreID & ", "
         Else
            vSQL = vSQL & vbCrLf _
            & GridTable.Columns("Column_Name").Value & " = T2." & GridTable.Columns("Column_Name").Value & ", "
         End If
      End If
      GridTable.MoveNext
   Next i
   GridTable.MoveFirst
   
   vSQL = Replace(vSQL, "''", "Null")
   vSQL = Left(vSQL, Len(vSQL) - 2)
   vSQL = vSQL & vbCrLf _
   & "From JournalVouchersBody B " & vbCrLf _
   & "inner join JournalVouchers H on H.SID = B.SID" & vbCrLf _
   & "Inner Join " & vbCrLf _
   & "(" & vbCrLf _
   & " Select " & vColumnList2 & vbCrLf _
   & " from " & vConfig & "JournalVouchersBody T2 " & vbCrLf _
   & " inner join " & vConfig & "JournalVouchers T1 " & vbCrLf _
   & " on T1.SID = T2.SID " & vbCrLf _
   & " inner join  JournalVouchers H on H.Tag = T1.SID And H.StoreID = " & vStoreID & vbCrLf _
   & " Where T1.Tag is Null And T1.IsSync = 0 And H.StoreID = " & vStoreID & vbCrLf _
   & " )T2"
   vSQL = vSQL & vbCrLf _
   & " on H.Tag = T2.SID And T2.AccountNo = B.AccountNo And H.StoreID = " & vStoreID & vbCrLf _
   & " Where H.Tag is not Null "
   CN.Execute vSQL

   ''''''' Insert record
          
   vColumnList2 = Replace(UCase(vColumnList2), "T2.SID", "H.SID")
   vColumnList2 = Replace(UCase(vColumnList2), UCase("T2.StoreID"), vStoreID)
   
   vSQL = "Insert into " & " JournalVouchersBody " & "(" & vColumnList1 & ")" & vbCrLf _
         & " Select " & vColumnList2 & vbCrLf _
         & " from " & vConfig & "JournalVouchersBody T2" & vbCrLf _
         & " inner join " & vConfig & "JournalVouchers T1" & vbCrLf _
         & " on T2.SID = T1.SID " & vbCrLf _
         & " inner join  JournalVouchers H on H.Tag = T1.SID And H.StoreID =" & vStoreID & vbCrLf _
         & " left outer Join JournalVouchersBody B on B.SID = H.SID And B.AccountNo = T2.AccountNo" & vbCrLf _
         & " Where T1.IsSync = 0 And B.SerialNo is Null And (T1.Tag is Null or T1.Tag = '') And H.StoreID =" & vStoreID
      
   CN.Execute vSQL
     
   ''''''''''''' Update Remotly Client System '''''''''''''''''''''
   vSQL = "Update " & vConfig & "JournalVouchers Set IsSync =1 where IsSync = 0"
   CN.Execute vSQL
   
   ''''''''''''' Delete Journal Body '''''''''''''''''''''
   vSQL = "Delete JournalVouchersBody " & vbCrLf _
          & " From JournalVouchersBody b" & vbCrLf _
          & " Inner join JournalVouchers H on H.SID = B.SID " & vbCrLf _
          & " Left outer join " & vConfig & "JournalVouchers T1 on T1.SID = H.Tag And H.StoreID = " & vStoreID & vbCrLf _
          & " Left outer join " & vConfig & "JournalVouchersBody T2 on T1.SID = T2.SID And T1.SID = H.Tag And T2.AccountNo = b.AccountNo" & vbCrLf _
          & " WHere H.Tag is Not Null And T2.SerialNo is null And H.StoreID = " & vStoreID
  CN.Execute vSQL
  ''''''''''''' Delete Journal JournalVouchers '''''''''''''''''''''
  vSQL = "Delete JournalVouchers " & vbCrLf _
          & " From JournalVouchers H" & vbCrLf _
          & " Left outer join " & vConfig & "JournalVouchers T1 on T1.SID = H.Tag And H.StoreID = " & vStoreID & vbCrLf _
          & " WHere H.Tag is Not Null And T1.SID is null And H.StoreID = " & vStoreID
   CN.Execute vSQL
Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
   
End Sub

Private Sub BtnSyncAll_Click()
 Call BtnDefinationExport_Click
 Call BtnTransactionImport_Click
 Call BtnStockTrnsfer_Click
End Sub
