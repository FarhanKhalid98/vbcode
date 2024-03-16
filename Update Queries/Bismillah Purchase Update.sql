--Purchase Return Header
EXEC sp_rename 'PurchaseReturnHeader.Discount', 'BillDiscount', 'COLUMN'
go
ALTER TABLE PurchaseReturnHeader ADD StoreID tinyint NULL
go
update PurchaseReturnHeader set StoreID=1
go
ALTER TABLE PurchaseReturnHeader ALTER COLUMN StoreID tinyint not null
go

--Purchase Return Body
EXEC sp_rename 'PurchaseReturnBody.DiscountValue', 'DiscPC', 'COLUMN'
go
EXEC sp_rename 'PurchaseReturnBody.Qty', 'QtyLoose', 'COLUMN'
go
ALTER TABLE PurchaseReturnBody ALTER COLUMN DiscPC numeric(5,2) not NULL
go

ALTER TABLE PurchaseReturnBody ADD Code varchar(50) NULL
go
update PurchaseReturnBody set code=ProductID
go
ALTER TABLE PurchaseReturnBody ALTER COLUMN Code varchar(50) not null
go

ALTER TABLE PurchaseReturnBody ADD DiscPer numeric(4,2) NULL
go
update PurchaseReturnBody set DiscPer= Round((DiscPC*100)/Price, 2)
go
ALTER TABLE PurchaseReturnBody ALTER COLUMN DiscPer numeric(4,2) not NULL
go

ALTER TABLE PurchaseReturnBody ADD DiscVal numeric(7,2) NULL
go
UPDATE PurchaseReturnBody set DiscVal=QtyLoose*DiscPC
go
ALTER TABLE PurchaseReturnBody ALTER COLUMN DiscVal numeric(7,2) not NULL
go

ALTER TABLE PurchaseReturnBody ADD PackingID tinyint NULL
go
ALTER TABLE PurchaseReturnBody ADD QtyPack numeric(5,0) NULL
go
ALTER TABLE PurchaseReturnBody ADD Multiplier smallint NULL
go

--Purchase Header
EXEC sp_rename 'PurchaseHeader.Discount', 'BillDiscount', 'COLUMN'
go
ALTER TABLE PurchaseHeader ADD StoreID tinyint NULL
go
update PurchaseHeader set StoreID=1
go
ALTER TABLE PurchaseHeader ALTER COLUMN StoreID tinyint not null
go
ALTER TABLE PurchaseHeader ADD EntryDate smalldatetime NULL
go


--Purchase Body
EXEC sp_rename 'PurchaseBody.DiscountValue', 'DiscPC', 'COLUMN'
go
EXEC sp_rename 'PurchaseBody.Qty', 'QtyLoose', 'COLUMN'
go
ALTER TABLE PurchaseBody ALTER COLUMN DiscPC numeric(5,2) not NULL
go

ALTER TABLE PurchaseBody ADD Code varchar(50) NULL
go
update PurchaseBody set code=ProductID
go
ALTER TABLE PurchaseBody ALTER COLUMN Code varchar(50) not null
go

ALTER TABLE PurchaseBody ADD DiscPer numeric(4,2) NULL
go
update PurchaseBody set DiscPer= Round((DiscPC*100)/Price, 2)
go
ALTER TABLE PurchaseBody ALTER COLUMN DiscPer numeric(4,2) not NULL
go

ALTER TABLE PurchaseBody ADD DiscVal numeric(7,2) NULL
go
UPDATE PurchaseBody set DiscVal=QtyLoose*DiscPC
go
ALTER TABLE PurchaseBody ALTER COLUMN DiscVal numeric(7,2) not NULL
go

ALTER TABLE PurchaseBody ADD PackingID tinyint NULL
go
ALTER TABLE PurchaseBody ADD QtyPack numeric(5,0) NULL
go
ALTER TABLE PurchaseBody ADD Multiplier smallint NULL
go
