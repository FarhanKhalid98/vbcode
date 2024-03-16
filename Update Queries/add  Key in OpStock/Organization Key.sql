ALTER TABLE [dbo].[OpeningStock] DROP CONSTRAINT [PK_OpeningStock]
GO 
update OpeningStock set OrganizationID = 1
GO
ALTER TABLE OpeningStock ALTER COLUMN OrganizationID tinyint not null
GO
ALTER TABLE [dbo].[OpeningStock] WITH NOCHECK ADD 
	CONSTRAINT [PK_OpeningStock] PRIMARY KEY  CLUSTERED 
	(
		[ProductID],
		[StoreID],
		[OrganizationID]
	)  ON [PRIMARY] 
GO 

