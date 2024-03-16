sp_rename 'Users.IsEditDelete', 'IsEdit', 'COLUMN'
go
update users set IsEdit = 1
go
ALTER TABLE Users ALTER COLUMN IsEdit bit not NULL
go
ALTER TABLE Users ADD IsDelete bit NULL
go
update users set IsDelete = 1
go
ALTER TABLE Users ALTER COLUMN IsDelete bit not NULL
go
------------------------------------

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Tasks]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Tasks]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UserTasks]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UserTasks]
GO

CREATE TABLE [dbo].[Tasks] (
	[TaskKey] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TaskGroup] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserTasks] (
	[UserNo] [tinyint] NOT NULL ,
	[TaskKey] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Registry]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Registry]
GO

CREATE TABLE [dbo].[Registry] (
	[StoreID] [tinyint] NOT NULL ,
	[StoreVisible] [bit] NOT NULL ,
	[AddSpace] [bit] NOT NULL ,
	[NegativeSale] [bit] NOT NULL ,
	[BankMachineID] [tinyint] NULL ,
	[CashReceived] [bit] NOT NULL,
	[DuplicateCode] [bit] NOT NULL
) ON [PRIMARY]
GO
insert into [Registry] ([StoreID],StoreVisible,AddSpace,NegativeSale,BankMachineID,CashReceived,DuplicateCode) values (1,1,1,1,null,0,1)
go
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Stores]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Stores]
GO

CREATE TABLE [dbo].[Stores] (
	[StoreID] [tinyint] NOT NULL ,
	[StoreName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Stores] WITH NOCHECK ADD 
	CONSTRAINT [PK_Stores] PRIMARY KEY  CLUSTERED 
	(
		[StoreID]
	)  ON [PRIMARY] 
GO
insert into [Stores] (StoreID,StoreName) values (1,'Main')
go

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Company]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Company]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Manufacturer]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Manufacturer]
GO

CREATE TABLE [dbo].[Company] (
	[CompanyName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Address] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[City] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PhoneNo] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EMail] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Manufacturer] (
	[Name] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO
insert into [Manufacturer] ([Name]) values ('Developed by '' S o f t   I n n '' contact us at  0333-6134224')
go
insert into [Company] ([CompanyName],Address,City,PhoneNo,EMail) values ('786 Self Store','Ehle-Hadees Chowk','Khanewal','065-2558457, 0300-6890102','')
go

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CurrentStock]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CurrentStock]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CurrentStockExpiry]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CurrentStockExpiry]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CurrentStockStore]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CurrentStockStore]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CurrentStockWastage]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CurrentStockWastage]
GO

CREATE TABLE [dbo].[CurrentStock] (
	[ProductID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[QtyLoose] [numeric](9, 3) NOT NULL ,
	[Cost] [numeric](9, 3) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CurrentStockExpiry] (
	[ProductID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EQtyLoose] [numeric](9, 3) NOT NULL ,
	[DQtyLoose] [numeric](9, 3) NOT NULL ,
	[Cost] [decimal](18, 6) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CurrentStockStore] (
	[ProductID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[StoreID] [tinyint] NOT NULL ,
	[QtyLoose] [numeric](9, 3) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CurrentStockWastage] (
	[ProductID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[QtyLoose] [numeric](9, 3) NOT NULL ,
	[Cost] [decimal](18, 6) NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[CurrentStock] WITH NOCHECK ADD 
	CONSTRAINT [PK_CurrentStockCost] PRIMARY KEY  CLUSTERED 
	(
		[ProductID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CurrentStockExpiry] WITH NOCHECK ADD 
	CONSTRAINT [PK_CurrentStockExpiry] PRIMARY KEY  CLUSTERED 
	(
		[ProductID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CurrentStockStore] WITH NOCHECK ADD 
	CONSTRAINT [PK_CurrentStock] PRIMARY KEY  CLUSTERED 
	(
		[ProductID],
		[StoreID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CurrentStockWastage] WITH NOCHECK ADD 
	CONSTRAINT [PK_CurrentWastage] PRIMARY KEY  CLUSTERED 
	(
		[ProductID]
	)  ON [PRIMARY] 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AccountsBalances]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AccountsBalances]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BankMachines]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BankMachines]
GO

CREATE TABLE [dbo].[AccountsBalances] (
	[AccountNo] [varchar] (11) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OpeningDebit] [numeric](12, 2) NULL ,
	[OpeningCredit] [numeric](12, 2) NULL ,
	[OpeningBal] [numeric](12, 2) NULL ,
	[OpeningBalType] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Debit] [numeric](12, 2) NULL ,
	[Credit] [numeric](12, 2) NULL ,
	[Bal] [numeric](12, 2) NULL ,
	[BalType] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BankMachines] (
	[BankMachineID] [tinyint] NOT NULL ,
	[BankMachineName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AccountNo] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[BankMachines] WITH NOCHECK ADD 
	CONSTRAINT [PK_BankMachines] PRIMARY KEY  CLUSTERED 
	(
		[BankMachineID]
	)  ON [PRIMARY] 
GO

