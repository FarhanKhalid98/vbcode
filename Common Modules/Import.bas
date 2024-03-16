Attribute VB_Name = "Import"
Option Explicit

Private Sub SubGroups()
On Error GoTo ErrorHandler
'step - 1 add new rows
vStr = " insert into subgroups" & vbCrLf _
      + " select a.Subgroupid, a.SubgroupName " & vbCrLf _
      + " from subgroups s right outer join " & vbCrLf _
      + " OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls', [SubGroups$]) a " & vbCrLf _
      + " on a.subgroupid = s.subgroupid " & vbCrLf _
      + " where s.subgroupid is null"
   CN.Execute vStr
'step - 2 Copy into Temp DB
vStr = " insert into tempdb..subgroups" & vbCrLf _
      + " select a.SubGroupID, a.SubGroupName" & vbCrLf _
      + " from subgroups s inner join " & vbCrLf _
      + " OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls', [SubGroups$]) a " & vbCrLf _
      + " on a.subgroupid = s.subgroupid "
   CN.Execute vStr

'step - 3 Update Current DB with Temp DB
vStr = " update subgroups set SubgroupName = a.SubgroupName" & vbCrLf _
      + " from subgroups s inner join " & vbCrLf _
      + " tempdb..subgroups a " & vbCrLf _
      + " on a.subgroupid = s.subgroupid "
   CN.Execute vStr
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


Private Sub Groups()
On Error GoTo ErrorHandler
'step - 1 add new rows
vStr = " insert into groups" & vbCrLf _
      + " select a.groupid, a.GroupName " & vbCrLf _
      + " from Groups s right outer join " & vbCrLf _
      + " OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls', [Groups$]) a " & vbCrLf _
      + " on a.Groupid = s.Groupid " & vbCrLf _
      + " where s.Groupid is null"
   CN.Execute vStr
'step - 2 Copy into Temp DB
vStr = " insert into tempdb..Groups" & vbCrLf _
      + " select a.GroupID, a.GroupName" & vbCrLf _
      + " from Groups s inner join " & vbCrLf _
      + " OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls', [Groups$]) a " & vbCrLf _
      + " on a.Groupid = s.Groupid "
   CN.Execute vStr

'step - 3 Update Current DB with Temp DB
vStr = " update Groups set GroupName = a.GroupName" & vbCrLf _
      + " from Groups s inner join " & vbCrLf _
      + " tempdb..Groups a " & vbCrLf _
      + " on a.Groupid = s.Groupid "
   CN.Execute vStr
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Companies()
On Error GoTo ErrorHandler
'step - 1 add new rows
vStr = " insert into Companies" & vbCrLf _
      + " select a.Companyid, a.CompanyName " & vbCrLf _
      + " from Companies s right outer join " & vbCrLf _
      + " OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls', [Companies$]) a " & vbCrLf _
      + " on a.Companyid = s.Companyid " & vbCrLf _
      + " where s.Companyid is null"
   CN.Execute vStr
'step - 2 Copy into Temp DB
vStr = " insert into tempdb..Companies" & vbCrLf _
      + " select a.CompanyID, a.CompanyName" & vbCrLf _
      + " from Companies s inner join " & vbCrLf _
      + " OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls', [Companies$]) a " & vbCrLf _
      + " on a.Companyid = s.Companyid "
   CN.Execute vStr

'step - 3 Update Current DB with Temp DB
vStr = " update Companies set CompanyName = a.CompanyName" & vbCrLf _
      + " from Companies s inner join " & vbCrLf _
      + " tempdb..Companies a " & vbCrLf _
      + " on a.Companyid = s.Companyid "
   CN.Execute vStr
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub
