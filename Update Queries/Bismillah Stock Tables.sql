if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExpiryBody]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ExpiryBody]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExpiryClaimsBody]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ExpiryClaimsBody]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExpiryClaimsHeader]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ExpiryClaimsHeader]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExpiryHeader]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ExpiryHeader]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExpiryReplyBody]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ExpiryReplyBody]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExpiryReplyHeader]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ExpiryReplyHeader]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ManufacturedProductsBody]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ManufacturedProductsBody]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ManufacturedProductsHeader]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ManufacturedProductsHeader]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ManufacturedProductsUsed]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ManufacturedProductsUsed]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ProductProcessInfoBody]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ProductProcessInfoBody]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ProductProcessInfoHeader]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ProductProcessInfoHeader]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StockTransferBody]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StockTransferBody]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StockTransferHeader]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StockTransferHeader]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StockWastageBody]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StockWastageBody]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StockWastageHeader]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StockWastageHeader]
GO

CREATE TABLE [dbo].[ExpiryBody] (
	[SerialNo] [int] IDENTITY (1, 1) NOT NULL ,
	[ExpiryID] [smallint] NOT NULL ,
	[ExpiryDate] [smalldatetime] NOT NULL ,
	[PackingID] [tinyint] NOT NULL ,
	[Code] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ProductID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EQtyPack] [numeric](9, 0) NOT NULL ,
	[EQtyLoose] [numeric](9, 3) NOT NULL ,
	[DQtyPack] [numeric](9, 0) NOT NULL ,
	[DQtyLoose] [numeric](9, 3) NOT NULL ,
	[Multiplier] [smallint] NOT NULL ,
	[Cost] [numeric](9, 2) NOT NULL ,
	[Amount] [numeric](18, 3) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ExpiryClaimsBody] (
	[ClaimId] [int] NOT NULL ,
	[ProductId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Code] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PackingId] [tinyint] NOT NULL ,
	[EQtyPack] [numeric](9, 0) NOT NULL ,
	[EQtyLoose] [numeric](9, 3) NOT NULL ,
	[DQtyPack] [numeric](9, 0) NOT NULL ,
	[DQtyLoose] [numeric](9, 3) NOT NULL ,
	[Multiplier] [smallint] NOT NULL ,
	[Cost] [numeric](9, 2) NOT NULL ,
	[Amount] [numeric](18, 3) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ExpiryClaimsHeader] (
	[ClaimId] [int] NOT NULL ,
	[ClaimDate] [smalldatetime] NOT NULL ,
	[TotalAmount] [numeric](18, 3) NOT NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[UserNo] [tinyint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ExpiryHeader] (
	[ExpiryID] [smallint] NOT NULL ,
	[ExpiryDate] [smalldatetime] NOT NULL ,
	[StoreID] [tinyint] NOT NULL ,
	[TotalAmount] [numeric](18, 3) NOT NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[UserNo] [tinyint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ExpiryReplyBody] (
	[ReplyId] [int] NOT NULL ,
	[ProductId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Code] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PackingId] [tinyint] NOT NULL ,
	[Multiplier] [smallint] NOT NULL ,
	[QtyPack] [numeric](9, 0) NOT NULL ,
	[QtyLoose] [numeric](9, 3) NOT NULL ,
	[Cost] [numeric](9, 2) NOT NULL ,
	[Amount] [numeric](18, 3) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ExpiryReplyHeader] (
	[ReplyId] [int] NOT NULL ,
	[ReplyDate] [smalldatetime] NOT NULL ,
	[StoreID] [tinyint] NOT NULL ,
	[TotalAmount] [numeric](18, 3) NOT NULL ,
	[RepliedAmount] [numeric](18, 3) NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[UserNo] [tinyint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ManufacturedProductsBody] (
	[ManufacturedID] [int] NOT NULL ,
	[Code] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ProductID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[QtyLoose] [numeric](9, 3) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ManufacturedProductsHeader] (
	[ManufacturedID] [int] NOT NULL ,
	[ManufacturedDate] [smalldatetime] NOT NULL ,
	[StoreID] [tinyint] NOT NULL ,
	[UserNo] [tinyint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ManufacturedProductsUsed] (
	[ManufacturedID] [int] NOT NULL ,
	[ProductID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[QtyLoose] [numeric](9, 3) NOT NULL ,
	[Rate] [numeric](9, 2) NOT NULL ,
	[Amount] [numeric](18, 3) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ProductProcessInfoBody] (
	[ID] [smallint] NOT NULL ,
	[Code] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ProductID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[QtyLoose] [numeric](9, 3) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ProductProcessInfoHeader] (
	[ID] [smallint] NOT NULL ,
	[FinishedProductID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[StockTransferBody] (
	[SerialNo] [bigint] IDENTITY (1, 1) NOT NULL ,
	[TransferID] [smallint] NOT NULL ,
	[TransferDate] [smalldatetime] NOT NULL ,
	[PackingID] [tinyint] NOT NULL ,
	[Code] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ProductID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[QtyPack] [numeric](9, 3) NOT NULL ,
	[QtyLoose] [numeric](8, 3) NOT NULL ,
	[Multiplier] [smallint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[StockTransferHeader] (
	[TransferID] [smallint] NOT NULL ,
	[TransferDate] [smalldatetime] NOT NULL ,
	[FromStoreID] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ToStoreID] [tinyint] NOT NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[UserNo] [tinyint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[StockWastageBody] (
	[SerialNo] [bigint] IDENTITY (1, 1) NOT NULL ,
	[WastageID] [smallint] NOT NULL ,
	[WastageDate] [smalldatetime] NOT NULL ,
	[PackingID] [tinyint] NOT NULL ,
	[Code] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ProductID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[QtyPack] [numeric](9, 3) NOT NULL ,
	[QtyLoose] [numeric](9, 3) NOT NULL ,
	[Multiplier] [smallint] NOT NULL ,
	[Cost] [numeric](9, 2) NOT NULL ,
	[Amount] [numeric](18, 3) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[StockWastageHeader] (
	[WastageID] [smallint] NOT NULL ,
	[WastageDate] [smalldatetime] NOT NULL ,
	[StoreID] [tinyint] NOT NULL ,
	[TotalAmount] [numeric](18, 3) NOT NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[UserNo] [tinyint] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ExpiryBody] WITH NOCHECK ADD 
	CONSTRAINT [PK_ExpiryBody] PRIMARY KEY  CLUSTERED 
	(
		[SerialNo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ExpiryClaimsHeader] WITH NOCHECK ADD 
	CONSTRAINT [PK_ExpiryClaimsHeader] PRIMARY KEY  CLUSTERED 
	(
		[ClaimId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ExpiryHeader] WITH NOCHECK ADD 
	CONSTRAINT [PK_ExpiryHeader] PRIMARY KEY  CLUSTERED 
	(
		[ExpiryID],
		[ExpiryDate]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ExpiryReplyHeader] WITH NOCHECK ADD 
	CONSTRAINT [PK_ExpiryClaimReplyHeader] PRIMARY KEY  CLUSTERED 
	(
		[ReplyId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ManufacturedProductsHeader] WITH NOCHECK ADD 
	CONSTRAINT [PK_ManufacturedProductsHeader] PRIMARY KEY  CLUSTERED 
	(
		[ManufacturedID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ProductProcessInfoHeader] WITH NOCHECK ADD 
	CONSTRAINT [PK_ProductProcessInfoHeader] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StockTransferBody] WITH NOCHECK ADD 
	CONSTRAINT [PK_StockTransferHeader] PRIMARY KEY  CLUSTERED 
	(
		[SerialNo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StockTransferHeader] WITH NOCHECK ADD 
	CONSTRAINT [PK_StockTraansferHeader] PRIMARY KEY  CLUSTERED 
	(
		[TransferID],
		[TransferDate]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StockWastageBody] WITH NOCHECK ADD 
	CONSTRAINT [PK_StockWastageBody] PRIMARY KEY  CLUSTERED 
	(
		[SerialNo]
	)  ON [PRIMARY] 
GO

