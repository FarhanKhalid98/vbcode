ALTER TABLE SaleHeader ADD OrderID SmallInt Null 
GO
ALTER TABLE SaleHeader ADD OrderDate SmallDateTime Null 
GO
ALTER TABLE PurchaseHeader ADD OrderID SmallInt Null 
GO
ALTER TABLE PurchaseHeader ADD OrderDate SmallDateTime Null 
GO
ALTER TABLE SaleOrderHeader ADD IsSale Bit Not Null Default 0
GO
ALTER TABLE PurchaseOrderHeader ADD IsPurchase Bit Not Null Default 0
GO