CREATE TRIGGER [tu_purchasesbody] ON dbo.PurchaseBody 
FOR UPDATE
AS
BEGIN

DECLARE @ErrorMsg as varchar(100)

If (@@ROWCOUNT>1)
BEGIN
	set @ErrorMsg= 'Multiple rows cannot be updated.'
	goto error
END
/*************CURRENT STOCK QTY*******************************/
DECLARE @CSQty as numeric(9,3),@CSCost as numeric(9,3),  @CSTTLCost as numeric(18,6)
/*************CURRENT STOCK STORE QTY*******************************/
DECLARE @CSSQty as numeric(9,3)
/***** Declaring variables to be inserted *******/
DECLARE @IProductId as varchar(5), @IStoreID as tinyint, @IQty as numeric(9,3), @ICost as numeric(9,3),
@NewTTLCost as numeric(18,6)
/***** Declaring variables to be Deleted *******/
DECLARE @DProductId as varchar(5), @DStoreID as tinyint, @DQty as numeric(9,3), @DCost as numeric(9,3),
@DelTTLCost as numeric(18,6)
/**************************************************************************/
Select @IProductId=ProductId, @IStoreID=StoreID, @IQty=(isnull(QtyPack,0)*isnull(Multiplier,0))+QtyLoose , @ICost=Price-isnull(DiscPC,0), @NewTTLCost = Round(@IQty*@ICost, 6)
From INSERTED i inner join purchaseheader h on i.purid = h.purid  and i.purchasedate = h.purchasedate
/***************************************************/
Select @DProductId=ProductId, @DStoreId=StoreId, @DQty=(isnull(QtyPack,0)*isnull(Multiplier,0))+QtyLoose, @DCost=Price-isnull(DiscPC,0), @DelTTLCost = @DQty*@DCost
From DELETED d inner join purchaseheader h on d.purid = h.purid  and d.purchasedate = h.purchasedate
/***************************************************/
Select @CSQty=QtyLoose, @CSCost=Cost, @CSTTLCost=QtyLoose*Cost From CurrentStock Where ProductId=@DProductId
/***************************************************/
Select @CSSQty=QtyLoose From CurrentStockStore Where ProductId=@DProductId and StoreID=@DStoreID
/***************************************************/
If not exists (Select ProductId From CurrentStockStore Where ProductId=@IProductId and StoreID=@IStoreID)
  Begin
	Insert Into CurrentStockStore(ProductId, StoreID, QtyLoose) Values(@IProductId, @IStoreId, @IQty-@DQty)	
  end
else
  begin
	Update CurrentStockStore Set QtyLoose=QtyLoose-@DQty+@IQty
	Where ProductId=@IProductId and StoreID=@IStoreID						
  end
If not exists (Select ProductId From CurrentStock Where ProductId=@IProductId)
  begin
	Insert Into CurrentStock(ProductId, QtyLoose, Cost) Values(@IProductId, @IQty-@DQty, @ICost)				
  end
else
  begin
	If (@CSQty - @DQty + @IQty = 0)
	  begin
		Update CurrentStock Set QtyLoose=QtyLoose-@DQty+@IQty, 
		Cost=(0)
		Where ProductId=@IProductId										
	  end
	else
	  begin
		Update CurrentStock Set QtyLoose=QtyLoose-@DQty+@IQty, 
		Cost=(@CSTTLCost - @DelTTLCost + @NewTTLCost) / (QtyLoose - @DQty + @IQty)
		Where ProductId=@IProductId						
	  end
  end
/***************************************************/
return
ERROR:
	raiserror(@ErrorMsg,16,1)
END



