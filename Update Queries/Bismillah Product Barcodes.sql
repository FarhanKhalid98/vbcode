if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ProductBarcodes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ProductBarcodes]
GO

CREATE TABLE [dbo].[ProductBarcodes] (
	[ProductID] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Code] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS not NULL 
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Products] DROP CONSTRAINT [PK_Products]
GO 
ALTER TABLE Products ADD PID varchar(5) NULL
go

-- Declare the variables to store the values returned by FETCH.
DECLARE @Code varchar(50), @ProductID varchar(7)
DECLARE Product_cursor CURSOR FOR
SELECT Code, ProductID FROM Products
where code is not null and code <> ''
ORDER BY ProductID

OPEN Product_cursor


FETCH NEXT FROM Product_cursor
INTO @Code, @ProductID

-- Check @@FETCH_STATUS to see if there are any more rows to fetch.
WHILE @@FETCH_STATUS = 0
BEGIN

	insert into ProductBarcodes (ProductID,Code) values( @ProductID,@Code)
   -- Concatenate and display the current values in the variables.

   -- This is executed as long as the previous fetch succeeds.
FETCH NEXT FROM Product_cursor
INTO @Code, @ProductID
END

CLOSE Product_cursor
DEALLOCATE Product_cursor
GO

-- select * from ProductBarcodes
-- delete from ProductBarcodes

-- Declare the variables to store the values returned by FETCH.
DECLARE @Code varchar(50), @ProductID varchar(7), @PID varchar(5)
DECLARE Product_cursor CURSOR FOR
SELECT Code, ProductID FROM Products
ORDER BY ProductID
OPEN Product_cursor

FETCH NEXT FROM Product_cursor
INTO @Code, @ProductID
set @PID = '00001'
-- Check @@FETCH_STATUS to see if there are any more rows to fetch.
WHILE @@FETCH_STATUS = 0
BEGIN

	update Products set PID = @PID where ProductID = @ProductID
	insert into ProductBarcodes (ProductID,Code) values( @PID,@ProductID)

   -- This is executed as long as the previous fetch succeeds.
FETCH NEXT FROM Product_cursor
INTO @Code, @ProductID

set @PID = right('00000' + cast(isnull(max(@PID),0) + 1 as varchar) ,5)
END

CLOSE Product_cursor
DEALLOCATE Product_cursor
GO

-- select * from ProductBarcodes


-- Declare the variables to store the values returned by FETCH.
DECLARE @Code varchar(50), @ProductID varchar(7), @PID varchar(5)
DECLARE Product_cursor CURSOR FOR
SELECT PID, ProductID FROM Products
ORDER BY ProductID
OPEN Product_cursor

FETCH NEXT FROM Product_cursor
INTO @PID, @ProductID
-- Check @@FETCH_STATUS to see if there are any more rows to fetch.
WHILE @@FETCH_STATUS = 0
BEGIN
	update productBarcodes set ProductID = @PID where ProductID = @ProductID

   -- This is executed as long as the previous fetch succeeds.
FETCH NEXT FROM Product_cursor
INTO @PID, @ProductID

END

CLOSE Product_cursor
DEALLOCATE Product_cursor
GO


ALTER TABLE Products DROP COLUMN code
go
ALTER TABLE Products DROP COLUMN ProductID
go
EXEC sp_rename 'Products.PID', 'ProductID', 'COLUMN'
go
ALTER TABLE Products ALTER COLUMN ProductID varchar(5) not null
go
ALTER TABLE [dbo].[Products] WITH NOCHECK ADD 
	CONSTRAINT [PK_Products] PRIMARY KEY  CLUSTERED 
	(
		[ProductID]
	)  ON [PRIMARY] 
GO 


-- Declare the variables to store the values returned by FETCH.
DECLARE @Code varchar(50), @ProductID varchar(7), @PID varchar(5)
DECLARE Product_cursor CURSOR FOR
select Code, ProductID from ProductBarcodes where len(code) = 7
ORDER BY Code
OPEN Product_cursor

FETCH NEXT FROM Product_cursor
INTO @Code, @ProductID
-- Check @@FETCH_STATUS to see if there are any more rows to fetch.
WHILE @@FETCH_STATUS = 0
BEGIN

	update Purchasebody set ProductID = @ProductID where ProductID = @Code
	update PurchaseReturnBody set ProductID = @ProductID where ProductID = @Code
	update SaleBody set ProductID = @ProductID where ProductID = @Code
	update SaleReturnBody set ProductID = @ProductID where ProductID = @Code
	
   -- This is executed as long as the previous fetch succeeds.
FETCH NEXT FROM Product_cursor
INTO @Code, @ProductID

END

CLOSE Product_cursor
DEALLOCATE Product_cursor
GO

ALTER TABLE Purchasebody ALTER COLUMN ProductID varchar(5) not null
go
ALTER TABLE PurchaseReturnBody ALTER COLUMN ProductID varchar(5) not null
go
ALTER TABLE SaleBody ALTER COLUMN ProductID varchar(5) not null
go
ALTER TABLE SaleReturnBody ALTER COLUMN ProductID varchar(5) not null
go