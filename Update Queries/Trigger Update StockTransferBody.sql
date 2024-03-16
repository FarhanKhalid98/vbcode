CREATE TRIGGER [tu_StockTransferBody] ON dbo.StockTransferBody 
FOR UPDATE
AS
BEGIN
/***** Declaring variables to be inserted *******/
DECLARE @ErrorMsg as varchar(100)

If (@@ROWCOUNT>1)
BEGIN
	set @ErrorMsg= 'Multiple rows cannot be deleted.'
	goto error
END
/********************************/
DECLARE @CSSFQty as decimal(18,3), @CSSTQty as decimal(18,3)
/****************************************************/
DECLARE @IProductId as varchar(5), @IFStoreID as tinyint, @IQty as decimal(9,3), @ITStoreID as tinyint
/****************************************************/
DECLARE @DProductId as varchar(5), @DFStoreID as tinyint, @DQty as decimal(9,3), @DTStoreID as tinyint
/****************************************************/
Select @DProductId=ProductId, @DFStoreID=FromStoreID, @DTStoreID=ToStoreID, @DQty=(isnull(QtyPack,0)*isnull(Multiplier,0))+QtyLoose
From DELETED d inner join StockTransferHeader h on d.TransferID = h.TransferID and d.TransferDate = h.TransferDate
/***************************************************/
Select @IProductId=ProductId, @IFStoreID=FromStoreID, @ITStoreID=ToStoreID, @IQty=(isnull(QtyPack,0)*isnull(Multiplier,0))+QtyLoose
From INSERTED i inner join StockTransferHeader h on i.TransferID = h.TransferID and i.TransferDate = h.TransferDate
/***************************************************/
If not exists (Select ProductId From CurrentStockStore Where ProductId=@IProductId and StoreId =@IFStoreID)
	BEGIN
		Insert Into CurrentStockStore(ProductId, StoreID, QtyLoose) Values(@IProductId, @IFStoreId, @DQty - @IQty)						
	END
ELSE
	BEGIN
		Update CurrentStockStore Set QtyLoose=QtyLoose + @DQty - @IQty
		Where ProductId=@IProductId and StoreID=@IFStoreId
	END
If not exists (Select ProductId From CurrentStockStore Where ProductId=@IProductId and storeid =@ITStoreID)
	BEGIN
		Insert Into CurrentStockStore(ProductId, StoreID, QtyLoose) Values(@IProductId, @ITStoreId, @IQty - @DQty)
	END
ELSE
	BEGIN
		Update CurrentStockStore Set QtyLoose=QtyLoose - @DQty + @IQty
		Where ProductId=@IProductId and StoreID=@ITStoreId
	END

/***************************************************/
return
ERROR:
	raiserror(@ErrorMsg,16,1)
END