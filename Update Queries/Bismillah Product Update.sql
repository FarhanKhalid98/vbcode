--Product
ALTER TABLE Products ADD DiscPer numeric(6,3) NULL
go
ALTER TABLE Products ADD DiscPC numeric(6,3) NULL
go
ALTER TABLE Products ALTER COLUMN PurchasePackingID tinyint NULL
go
ALTER TABLE Products ALTER COLUMN SubGroupID smallint NULL
go
ALTER TABLE Products ALTER COLUMN CompanyID smallint NULL
go
ALTER TABLE Products ALTER COLUMN PurPrice numeric(9,3) not NULL
go
ALTER TABLE Products ALTER COLUMN RetailPrice numeric(9,3) not NULL
go

-- Product Packing
ALTER TABLE ProductPacking ALTER COLUMN Multiplier smallint not NULL
go
ALTER TABLE ProductPacking ALTER COLUMN ProductID varchar(5) not NULL
go

-- Companies 
ALTER TABLE [dbo].[Companies] DROP CONSTRAINT [PK_Companies]
go
ALTER TABLE Companies ALTER COLUMN CompanyID smallint not NULL
go
go
ALTER TABLE [dbo].[Companies] WITH NOCHECK ADD 
	CONSTRAINT [PK_Companies] PRIMARY KEY  CLUSTERED 
	(
		[CompanyID]
	)  ON [PRIMARY] 
GO 

-- Sub Groups
ALTER TABLE [dbo].[SubGroups] DROP CONSTRAINT [PK_SubGroups]
go
ALTER TABLE SubGroups ALTER COLUMN SubGroupID smallint not NULL
go
go
ALTER TABLE [dbo].[SubGroups] WITH NOCHECK ADD 
	CONSTRAINT [PK_SubGroups] PRIMARY KEY  CLUSTERED 
	(
		[SubGroupID]
	)  ON [PRIMARY] 
GO 