if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SchProduct]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SchProduct]
GO

CREATE TABLE [dbo].[SchProduct] (
	[ProductID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Code] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ProductName] [char] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DiscPrice] [decimal](9, 3) NOT NULL ,
	[RetailPrice] [decimal](9, 3) NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[SchProduct] WITH NOCHECK ADD 
	CONSTRAINT [PK_SchProduct] PRIMARY KEY  CLUSTERED 
	(
		[ProductID]
	)  ON [PRIMARY] 
GO

delete from products where productname is null
delete from SchProduct

insert into SchProduct
select p.productID, code, Productname, RetailPrice-isnull(DiscPC,0) as Discprice, RetailPrice
from products p inner join ( select * from ProductBarcodes where len(code)=7)b on p.productid = b.productid