ALTER TABLE Products ADD IsWSSaleTax bit Not Null Default 0
GO
ALTER TABLE Products ADD IsRetailSaleTax bit Not Null Default 0
GO
ALTER TABLE Products ADD IsWSDiscb4ST bit Not Null Default 0
GO
ALTER TABLE Products ADD SaleTaxPer Numeric(9,3) Null
GO
ALTER TABLE SaleBody ADD RetailPrice Numeric(9,2) Not Null Default 0
GO
ALTER TABLE SaleBody ADD IsWSSaleTax bit Not Null Default 0
GO
ALTER TABLE SaleBody ADD IsRetailSaleTax bit Not Null Default 0
GO
ALTER TABLE SaleBody ADD IsWSDiscb4ST bit Not Null Default 0
GO
ALTER TABLE PurchaseBody ADD RetailPrice Numeric(9,2) Not Null Default 0
GO
ALTER TABLE PurchaseBody ADD IsWSSaleTax bit Not Null Default 0
GO
ALTER TABLE PurchaseBody ADD IsRetailSaleTax bit Not Null Default 0
GO
ALTER TABLE PurchaseBody ADD IsWSDiscb4ST bit Not Null Default 0
GO