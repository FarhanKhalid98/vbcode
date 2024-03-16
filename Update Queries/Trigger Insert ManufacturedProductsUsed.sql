CREATE TRIGGER [ti_ManufacturedProductsUsed] ON dbo.ManufacturedProductsUsed 
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
DECLARE @IProductId as varchar(5), @IStoreID as tinyint, @IQty as numeric(9,3), @ICost as numeric(9,3),
@CurrentTTLCost as numeric(18,3), @NewTTLCost as numeric(18,3)
/****************************************************/
Select @IProductId=ProductId, @IStoreID=StoreID, @IQty=QtyLoose, @ICost=Rate, @NewTTLCost = @IQty*@ICost
From INSERTED i inner join ManufacturedProductsHeader h on i.ManufacturedID = h.ManufacturedID
/***************************************************/
Select @CurrentTTLCost=Isnull((QtyLoose*Cost),0) From CurrentStock Where ProductId=@IProductId
/*****************************************************************************************************************/
If Not Exists (Select ProductId From CurrentStock Where ProductId=@IProductId)
  BEGIN
	Insert Into CurrentStock(ProductId, QtyLoose, Cost) Values(@IProductId, -@IQty, @ICost)	
  END
ELSE 
  begin
	If Exists (Select ProductId From CurrentStock Where ProductId=@IProductId and QtyLoose-@IQty<>0)
	  begin 
		Update CurrentStock Set QtyLoose=QtyLoose-(@IQty), 
		Cost=(@CurrentTTLCost - @NewTTLCost) / (QtyLoose-@IQty)
		Where ProductId=@IProductId
	  end
	else
	  begin
		Update CurrentStock Set QtyLoose=QtyLoose-(@IQty), 
		Cost=0
		Where ProductId=@IProductId
	  end
  end
if not Exists (Select ProductId From CurrentStockStore Where ProductId=@IProductId and StoreID=@IStoreId)
  BEGIN
	Insert Into CurrentStockStore(ProductId, StoreID, QtyLoose) Values(@IProductId, @IStoreId, -@IQty)		
  END
else
  BEGIN
		Update CurrentStockStore Set QtyLoose=QtyLoose-(@IQty)
		Where ProductId=@IProductId and StoreID=@IStoreId
  END
/***************************************************/
return
ERROR:
	raiserror (@ErrorMsg,16,1)
END



