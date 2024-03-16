Create TRIGGER [ti_StockTransferBody] ON dbo.StockTransferBody 
FOR INSERT
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
DECLARE @IProductId as varchar(5), @FStoreID as tinyint, @Qty as decimal(18,3), @TStoreID as tinyint
/****************************************************/
Select @IProductId=ProductId, @FStoreID=FromStoreID, @TStoreID=ToStoreID, @Qty=(isnull(QtyPack,0)*isnull(Multiplier,0))+QtyLoose
From INSERTED i inner join StockTransferHeader h on i.TransferID = h.TransferID and i.TransferDate = h.TransferDate
/***************************************************/
If not exists (Select ProductId From CurrentStockStore Where ProductId=@IProductId and StoreId =@TStoreID)
	BEGIN
		Insert Into CurrentStockStore(ProductId, StoreID, QtyLoose) Values(@IProductId, @TStoreId, @Qty)						
	END
ELSE
	BEGIN
		Update CurrentStockStore Set QtyLoose=QtyLoose+(@Qty)
		Where ProductId=@IProductId and StoreID=@TStoreId
	END
If not exists (Select ProductId From CurrentStockStore Where ProductId=@IProductId and storeid =@FStoreID)
	BEGIN
		Insert Into CurrentStockStore(ProductId, StoreID, QtyLoose) Values(@IProductId, @FStoreId, -@Qty)
	END
ELSE
	BEGIN
		Update CurrentStockStore Set QtyLoose=QtyLoose-(@Qty)
		Where ProductId=@IProductId and StoreID=@FStoreId
	END
/***************************************************/
return
ERROR:
	raiserror (@ErrorMsg,16,1)
END