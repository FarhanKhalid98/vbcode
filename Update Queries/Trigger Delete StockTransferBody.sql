CREATE TRIGGER [td_StockTransferBody] ON dbo.StockTransferBody 
FOR DELETE
AS
BEGIN
/***** Declaring variables to be inserted *******/
DECLARE @ErrorMsg as varchar(100)

If (@@ROWCOUNT>1)
BEGIN
	set @ErrorMsg= 'Multiple rows cannot be deleted.'
	goto error
END
/****************************************************/
DECLARE @IProductId as varchar(5), @FStoreID as tinyint, @Qty as decimal(9,3), @TStoreID as tinyint
/****************************************************/
Select @IProductId=ProductId, @FStoreID=FromStoreID, @TStoreID=ToStoreID, @Qty=(isnull(QtyPack,0)*isnull(Multiplier,0))+QtyLoose
From DELETED d inner join StockTransferHeader h on d.TransferID = h.TransferID and d.TransferDate = h.TransferDate
/***************************************************/
If not exists (Select ProductId From CurrentStockStore Where ProductId=@IProductId and StoreId =@FStoreID)
	BEGIN
		Insert Into CurrentStockStore(ProductId, StoreID, QtyLoose) Values(@IProductId, @FStoreId, @Qty)						
	END
ELSE
	BEGIN
		Update CurrentStockStore Set QtyLoose=QtyLoose+(@Qty)
		Where ProductId=@IProductId and StoreID=@FStoreId
	END
If not exists (Select ProductId From CurrentStockStore Where ProductId=@IProductId and storeid =@TStoreID)
	BEGIN
		Insert Into CurrentStockStore(ProductId, StoreID, QtyLoose) Values(@IProductId, @TStoreId, -@Qty)
	END
ELSE
	BEGIN
		Update CurrentStockStore Set QtyLoose=QtyLoose-(@Qty)
		Where ProductId=@IProductId and StoreID=@TStoreId
	END
/***************************************************/
return
ERROR:
	raiserror (@ErrorMsg,16,1)
END