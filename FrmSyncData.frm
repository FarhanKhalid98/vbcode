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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7782
      TabIndex        =   0
      Top             =   5745
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
      MICON           =   "FrmSyncData.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSync 
      Height          =   420
      Left            =   6364
      TabIndex        =   2
      Top             =   5745
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Sync"
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
      MICON           =   "FrmSyncData.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin MSComctlLib.ListView LVDef 
      Height          =   3030
      Left            =   360
      TabIndex        =   3
      Tag             =   "C"
      ToolTipText     =   "Product Entry"
      Top             =   1305
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   5345
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
      Left            =   3945
      TabIndex        =   4
      Top             =   5745
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
      MICON           =   "FrmSyncData.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin MSComctlLib.ListView LVTrans 
      Height          =   3030
      Left            =   5040
      TabIndex        =   5
      Tag             =   "C"
      ToolTipText     =   "Product Entry"
      Top             =   1260
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   5345
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
      Height          =   3030
      Left            =   9720
      TabIndex        =   6
      Top             =   1395
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
      stylesets(0).Picture=   "FrmSyncData.frx":0054
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
      stylesets(1).Picture=   "FrmSyncData.frx":0070
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
      stylesets(2).Picture=   "FrmSyncData.frx":008C
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
      stylesets(3).Picture=   "FrmSyncData.frx":00A8
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
      _ExtentY        =   5345
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
      TabIndex        =   1
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
Dim i, vPKeyCount, vIdentityKey, vStoreID, vTableCount As Integer
Dim FunGetMaxID As Integer
Dim vPrice, vMultiplier, vAmount, vTotalAmount As Double
Dim Rs As New ADODB.Recordset
Dim vLinkedServer As String


Private Sub BtnClose_Click()
   Unload Me
End Sub

Private Sub BtnRefresh_Click()
   Call Settings
End Sub

Private Sub BtnSync_Click()
   On Error GoTo ErrorHandler

      Me.MousePointer = vbHourglass
      vSQL = "Select * from Stores where StoreID <> 1 and isLock = 0 and Config is not null"
      If Rs.State = adStateOpen Then Rs.Close
      Rs.Open vSQL, CN, adOpenStatic, adLockReadOnly
   
      While Not Rs.EOF
         vStoreID = Rs!StoreID
         vConfig = Rs!Config
         Call SyncDefination
'         Call SyncTransaction
   '      Call SyncStockTransfer
         Rs.MoveNext
      Wend
         MsgBox "Sync Succeed", vbInformation, Me.Caption
         Me.MousePointer = vbDefault
   Exit Sub
ErrorHandler:
   Me.MousePointer = vbDefault
   Call ShowErrorMessage
End Sub

Private Sub SyncDefination()

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
               vColumnList1 = vColumnList1 + GridTable.Columns("Column_Name").Value & ", "
               vColumnList2 = vColumnList2 + "T2." + GridTable.Columns("Column_Name").Value & ", "
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
   vSQL = "Select * from " & LVDef.ListItems(vTableCount).Text & " T1 inner join " & vConfig & LVDef.ListItems(vTableCount).Text & " T2" & vJoin & vbCrLf _
         & " Where t1.modified_on >  (Select max(modified_on)  modified_on from " & vConfig & LVDef.ListItems(vTableCount).Text & " )"
   
   With CN.Execute(vSQL)
      While Not .EOF
         vSQL = "Update " & vConfig & LVDef.ListItems(vTableCount).Text & " Set "
         GridTable.MoveFirst
         For i = 0 To GridTable.Rows
            If Not (vSerialNo = GridTable.Columns("Column_Name").Value Or vPKey1 = GridTable.Columns("Column_Name").Value Or vPKey2 = GridTable.Columns("Column_Name").Value Or vPKey3 = GridTable.Columns("Column_Name").Value) Then
               vSQL = vSQL & vbCrLf _
               & GridTable.Columns("Column_Name").Value & " = '" & IIf(.Fields(i).Value = "True", 1, IIf(.Fields(i).Value = "False", 0, .Fields(i).Value)) & "', "
            End If
            GridTable.MoveNext
         Next i
         vSQL = Replace(vSQL, "''", "Null")
         vSQL = Left(vSQL, Len(vSQL) - 2)
         Select Case vPKeyCount
            Case 1
               vwhere = " Where " & vPKey1 & " = '" & .Fields(vPKey1).Value & "'"
         End Select
         vSQL = vSQL & vwhere
         CN.Execute vSQL
         GridTable.MoveFirst
         .MoveNext
      Wend
   End With
   
         
   ''''''' Insert record
      If LVDef.ListItems(vTableCount).Text <> "Products" Then
         vSQL = "Insert into " & vConfig & LVDef.ListItems(vTableCount).Text & IIf(vIdentityKey = 1, "(" & vColumnList1 & ")", "") & vbCrLf _
               & " Select " & IIf(vIdentityKey = 1, vColumnList1, "T1.*") & " from " & LVDef.ListItems(vTableCount).Text & " T1 Left outer join  " & vConfig & LVDef.ListItems(vTableCount).Text & " T2"
         vSQL = vSQL & vbCrLf _
            & vJoin & vbCrLf _
            & "where T2." & vPKey1 & "  is null" '" Where T2.modified_on >=  (Select isnull(max(modified_on),'01-01-1900')  modified_on from " & LVDef.ListItems(vTableCount).Text & " )" & vbCrLf _

         CN.Execute vSQL
      Else
         vSQL = "Select T1.* from " & LVDef.ListItems(vTableCount).Text & " T1 Left outer join " & vConfig & LVDef.ListItems(vTableCount).Text & " T2 " & vbCrLf _
         & vJoin & vbCrLf _
         & "where T2." & vPKey1 & "  is null"
         With CN.Execute(vSQL)
            While Not .EOF
               vSQL = "Insert into " & vConfig & LVDef.ListItems(vTableCount).Text & IIf(vIdentityKey = 1, "(" & vColumnList1 & ")", "") & vbCrLf _
               & " Select " & IIf(vIdentityKey = 1, vColumnList1, "T1.*") & " from " & LVDef.ListItems(vTableCount).Text & " T1 Left outer join  " & vConfig & LVDef.ListItems(vTableCount).Text & " T2"
                  vSQL = vSQL & vbCrLf _
                  & vJoin
               Select Case vPKeyCount
                  Case 1
                  vwhere = " Where T1." & vPKey1 & " = '" & .Fields(vPKey1).Value & "'"
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
      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "Colours")
      Item.SubItems(1) = "Colours"
      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "Sizes")
      Item.SubItems(1) = "Sizes"
      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "Seasons")
      Item.SubItems(1) = "Seasons"
      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "Descriptions")
      Item.SubItems(1) = "Descriptions"
      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "ItemDescription")
      Item.SubItems(1) = "ItemDescription"
      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "Parties")
      Item.SubItems(1) = "Parties "
      Item.Checked = True
      
'      Set Item = LVDef.ListItems.Add(, , "Users")
'      Item.SubItems(1) = "Users"
'      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "Products")
      Item.SubItems(1) = "Products"
      Item.Checked = True
                          
      Set Item = LVDef.ListItems.Add(, , "ChartOfAccounts")
      Item.SubItems(1) = "ChartOfAccounts"
      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "Designations")
      Item.SubItems(1) = "Designations"
      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "EmpDepartments")
      Item.SubItems(1) = "EmpDepartments"
      Item.Checked = True

      Set Item = LVDef.ListItems.Add(, , "Employees")
      Item.SubItems(1) = "Employees"
      Item.Checked = True
         
      Set Item = LVDef.ListItems.Add(, , "Members")
      Item.SubItems(1) = "Members"
      Item.Checked = True
      
      Set Item = LVDef.ListItems.Add(, , "MemberTypes")
      Item.SubItems(1) = "MemberTypes"
      Item.Checked = True
      
      Set Item = LVDef.ListItems.Add(, , "MembersDiscount")
      Item.SubItems(1) = "MembersDiscount"
      Item.Checked = True
      
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
      Item.Checked = True

'
'      Set Item = LVTrans.ListItems.Add(, , "SaleBody")
'      Item.SubItems(1) = "SaleHeader"
'      Item.Checked = False
'
'      Set Item = LVTrans.ListItems.Add(, , "SaleReturnHeader")
'      Item.SubItems(1) = "SaleReturnHeader"
'      Item.Checked = False
'
'      Set Item = LVTrans.ListItems.Add(, , "SaleReturnBody")
'      Item.SubItems(1) = "SaleReturnHeader"
'      Item.Checked = False
'
'      Set Item = LVTrans.ListItems.Add(, , "ReplacementHeader")
'      Item.SubItems(1) = "ReplacementHeader"
'      Item.Checked = False
'
'
'
'      Set Item = LVTrans.ListItems.Add(, , "DebitVouchers")
'      Item.SubItems(1) = "DebitVouchers"
'      Item.Checked = False
'
'      Set Item = LVTrans.ListItems.Add(, , "DebitVouchersBody")
'      Item.SubItems(1) = "DebitVouchers"
'      Item.Checked = False
'
'      Set Item = LVTrans.ListItems.Add(, , "CreditVouchers")
'      Item.SubItems(1) = "CreditVouchers"
'      Item.Checked = False
'
'      Set Item = LVTrans.ListItems.Add(, , "CreditVouchersBody")
'      Item.SubItems(1) = "CreditVouchers"
'      Item.Checked = False
'
'      Set Item = LVTrans.ListItems.Add(, , "JournalVouchers")
'      Item.SubItems(1) = "JournalVouchers"
'      Item.Checked = False
'
'      Set Item = LVTrans.ListItems.Add(, , "JournalVouchersBody")
'      Item.SubItems(1) = "JournalVouchers"
'      Item.Checked = False
'
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
Private Sub SyncTransaction()

'   If CN.State = adStateOpen Then CN.Close
'   CN.Open "Provider=SQLOLEDB.1;User ID=sa; Initial Catalog=" & vConnStr
'   CN.CursorLocation = adUseClient
'   CN.CommandTimeout = 200
   
   
      For vTableCount = 1 To LVTrans.ListItems.Count
      
         If LVTrans.ListItems(vTableCount).Checked = True Then
              
            vHeaderTable = LVTrans.ListItems(vTableCount).SubItems(1)
            
            Select Case LVTrans.ListItems(vTableCount).Text
               Case "Sale"
                  Call SaleTransaction
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
   & " inner join [113.203.194.101,1433].Test.dbo.SaleHeader T1 " & vbCrLf _
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
   
End Sub
