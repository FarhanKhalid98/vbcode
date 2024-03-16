Attribute VB_Name = "ImportExport"
Option Explicit

Public Sub Import()
   Call SubGroups
   Call Groups
   Call Companies
   Call Products
End Sub

Public Sub Export()
   On Error GoTo ErrorHandler
   ' subgroup
   vStr = " INSERT INTO " & vbCrLf _
      + " OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls;'," & vbCrLf _
      + " 'SELECT subgroupId, SubGroupName FROM [SubGroups$]') " & vbCrLf _
      + " SELECT SubGroupID, SubGroupName FROM SubGroups"
      CN.Execute vStr
   ' group
   vStr = " INSERT INTO " & vbCrLf _
      + " OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls;'," & vbCrLf _
      + " 'SELECT GroupId, GroupName FROM [Groups$]') " & vbCrLf _
      + " SELECT GroupID, GroupName FROM Groups"
      CN.Execute vStr
   ' Companies
   vStr = " INSERT INTO " & vbCrLf _
      + " OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls;'," & vbCrLf _
      + " 'SELECT CompanyId, CompanyName FROM [Companies$]') " & vbCrLf _
      + " SELECT CompanyID, CompanyName FROM Companies"
   CN.Execute vStr
   ' Products
   vStr = " INSERT INTO " & vbCrLf _
      + " OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls;'," & vbCrLf _
      + " 'SELECT ProductID, CompanyID, GroupID, SubGroupID, Productname, PurPrice," & vbCrLf _
      + " RetailPrice, PurchasePackingID, DiscPer, DiscPc, StockLimit, UnitID FROM [Products$]') " & vbCrLf _
      + " SELECT ProductID, CompanyID, GroupID, SubGroupID, Productname, PurPrice," & vbCrLf _
      + " RetailPrice, PurchasePackingID, DiscPer, DiscPc, StockLimit, UnitID  FROM Products"
   CN.Execute vStr
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


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

Private Sub Products()
On Error GoTo ErrorHandler
   CN.CommandTimeout = 1000
   CN.Execute "ALTER TABLE Products DISABLE TRIGGER [ti_Products]"
'step - 1 add new rows
vStr = " insert into Products (ProductID, CompanyID, GroupID, SubGroupID, Productname, PurPrice," & vbCrLf _
      + " RetailPrice, PurchasePackingID, DiscPer, DiscPc, StockLimit, UnitID )" & vbCrLf _
      + " select a.ProductID, a.CompanyID, a.GroupID, a.SubGroupID, a.Productname, a.PurPrice," & vbCrLf _
      + " a.RetailPrice, a.PurchasePackingID, a.DiscPer, a.DiscPc, a.StockLimit, a.UnitID " & vbCrLf _
      + " from Products s right outer join " & vbCrLf _
      + " OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls', [Products$]) a " & vbCrLf _
      + " on a.Productid = s.Productid " & vbCrLf _
      + " where s.Productid is null"
   CN.Execute vStr
   vStr = "insert into CurrentStock (ProductID, QtyLoose, Cost) select t.ProductID, 0 as QtyLoose, 0 as Cost from CurrentStock s right outer join OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls', [Products$]) t " & vbCrLf _
      + " on t.ProductID = s.ProductID where s.productid is null"
   CN.Execute vStr
   vStr = "insert into CurrentStockStore (ProductID, Storeid, QtyLoose) select t.ProductID, 1 as Storeid, 0 as QtyLoose from CurrentStockStore s right outer join OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls', [Products$]) t " & vbCrLf _
      + " on t.ProductID = s.ProductID where s.productid is null"
   CN.Execute vStr
   CN.Execute "ALTER TABLE Products ENABLE TRIGGER [ti_Products]"
   
'step - 2 Copy into Temp DB
vStr = " insert into tempdb..Products(ProductID, CompanyID, GroupID, SubGroupID, Productname, PurPrice," & vbCrLf _
      + " RetailPrice, PurchasePackingID, DiscPer, DiscPc, StockLimit, UnitID )" & vbCrLf _
      + " select a.ProductID, a.CompanyID, a.GroupID, a.SubGroupID, a.Productname, a.PurPrice," & vbCrLf _
      + " a.RetailPrice, a.PurchasePackingID, a.DiscPer, a.DiscPc, a.StockLimit, a.UnitID " & vbCrLf _
      + " from Products s inner join " & vbCrLf _
      + " OPENROWSET('Microsoft.Jet.OLEDB.4.0', " & vbCrLf _
      + " 'Excel 8.0;Database=" & App.Path & "\Data.xls', [Products$]) a " & vbCrLf _
      + " on a.Productid = s.Productid "
   CN.Execute vStr

'step - 3 Update Current DB with Temp DB
vStr = " update Products set CompanyID = a.CompanyID, GroupID = a.GroupID, SubGroupID = a.SubGroupID, Productname = a.Productname, " & vbCrLf _
      + " RetailPrice = a.RetailPrice, PurchasePackingID = a.PurchasePackingID, DiscPer = a.DiscPer, DiscPc = a.DiscPc, UnitID = a.UnitID, PurPrice = a.PurPrice " & vbCrLf _
      + " from Products s inner join " & vbCrLf _
      + " tempdb..Products a " & vbCrLf _
      + " on a.Productid = s.Productid "
   CN.Execute vStr
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub


