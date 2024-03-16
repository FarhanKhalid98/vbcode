--SaleReturnHeader
EXEC sp_rename 'SaleReturnHeader.Discount', 'BillDiscount', 'COLUMN'
go
EXEC sp_rename 'SaleReturnHeader.PaidAmount', 'CashPaid', 'COLUMN'
go
EXEC sp_rename 'SaleReturnHeader.TotalAmount', 'NetTotalAmount', 'COLUMN'
go
update SaleReturnHeader set CashPaid=NetTotalAmount
go
ALTER TABLE SaleReturnHeader ALTER COLUMN CashPaid numeric(5,0) not null
go
ALTER TABLE SaleReturnHeader ADD StoreID tinyint NULL
go
update SaleReturnHeader set StoreID=1
go
ALTER TABLE SaleReturnHeader ALTER COLUMN StoreID tinyint not null
go
ALTER TABLE SaleReturnHeader ADD CustomerName varchar(50) NULL
go
ALTER TABLE SaleReturnHeader ADD Credit bit NULL
go
update SaleReturnHeader set Credit=0
go
ALTER TABLE SaleReturnHeader ALTER COLUMN Credit bit not null
go
ALTER TABLE SaleReturnHeader ADD Cash bit NULL
go
update SaleReturnHeader set Cash=1
go
ALTER TABLE SaleReturnHeader ALTER COLUMN Cash bit not null
go

--Sale Return Body
EXEC sp_rename 'SaleReturnBody.DiscountValue', 'DiscPC', 'COLUMN'
go
ALTER TABLE SaleReturnBody ADD Code varchar(50) NULL
go
update SaleReturnBody set code=ProductID
go
ALTER TABLE SaleReturnBody ALTER COLUMN Code varchar(50) not null
go

ALTER TABLE SaleReturnBody ADD DiscPer numeric(4,2) NULL
go
update SaleReturnBody set DiscPer= Round((DiscPC*100)/Price, 2)
go
ALTER TABLE SaleReturnBody ALTER COLUMN DiscPer numeric(4,2) not NULL
go

ALTER TABLE SaleReturnBody ADD DiscVal numeric(7,2) NULL
go
update SaleReturnBody set DiscVal=Qty*DiscPC
go
ALTER TABLE SaleReturnBody ALTER COLUMN DiscVal numeric(7,2) not NULL
go

--Sale Header
EXEC sp_rename 'SaleHeader.SaleDate', 'BillDate', 'COLUMN'
go
EXEC sp_rename 'SaleHeader.Discount', 'BillDiscount', 'COLUMN'
go
EXEC sp_rename 'SaleHeader.TotalAmount', 'NetTotalAmount', 'COLUMN'
go
ALTER TABLE SaleHeader DROP COLUMN ReceivedAmount
go
ALTER TABLE SaleHeader ADD StoreID tinyint NULL
go
update SaleHeader set StoreID=1
go
ALTER TABLE SaleHeader ALTER COLUMN StoreID tinyint not null
go
ALTER TABLE SaleHeader ADD BankCard bit NULL
go
update SaleHeader set BankCard=0
go
ALTER TABLE SaleHeader ALTER COLUMN BankCard bit not null
go
ALTER TABLE SaleHeader ADD Credit bit NULL
go
update SaleHeader set Credit=0
go
ALTER TABLE SaleHeader ALTER COLUMN Credit bit not null
go
ALTER TABLE SaleHeader ADD Cash bit NULL
go
update SaleHeader set Cash=1
go
ALTER TABLE SaleHeader ALTER COLUMN Cash bit not null
go
ALTER TABLE SaleHeader ADD BankMachineID tinyint NULL
go
ALTER TABLE SaleHeader ADD InvoiceNo varchar(15) NULL
go
ALTER TABLE SaleHeader ADD CustomerName varchar(50) NULL
go


--Sale Body
EXEC sp_rename 'SaleBody.SaleDate', 'BillDate', 'COLUMN'
go
EXEC sp_rename 'SaleBody.DiscountValue', 'DiscPC', 'COLUMN'
go
ALTER TABLE SaleBody ADD Code varchar(50) NULL
go
update SaleBody set code=ProductID
go
ALTER TABLE SaleBody ALTER COLUMN Code varchar(50) not null
go

ALTER TABLE SaleBody ADD DiscPer numeric(4,2) NULL
go
update SaleBody set DiscPer= Round((DiscPC*100)/Price, 2)
go
ALTER TABLE SaleBody ALTER COLUMN DiscPer numeric(4,2) not NULL
go

ALTER TABLE SaleBody ADD DiscVal numeric(7,2) NULL
go
update SaleBody set DiscVal=Qty*DiscPC
go
ALTER TABLE SaleBody ALTER COLUMN DiscVal numeric(7,2) not NULL
go

ALTER TABLE SaleBody ADD Cost numeric(18,6) NULL
go
update SaleBody set Cost = 0
go
ALTER TABLE SaleBody ALTER COLUMN Cost numeric(9,3) not NULL
go