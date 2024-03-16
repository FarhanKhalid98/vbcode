CREATE TRIGGER [td_purchasesbody] ON dbo.PurchaseBody 
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
DECLARE @DProductId as varchar(5), @DStoreID as tinyint, @DQty as numeric(9,3), @DCost as numeric(9,3),
@CurrentTTLCost as numeric(18,3), @DelTTLCost as numeric(18,3)
/****************************************************/
Select @DProductId=ProductId, @DStoreId=StoreId, @DQty=(isnull(QtyPack,0)*isnull(Multiplier,0))+QtyLoose, @DCost=Price-isnull(DiscPC,0), @DelTTLCost = Round(@DQty * @DCost, 2)
From DELETED d inner join purchaseheader h on d.purid = h.purid  and d.purchasedate = h.purchasedate
/***************************************************/
Select @CurrentTTLCost=round(Isnull((QtyLoose*Cost),0),3) From CurrentStock Where ProductId=@DProductId
/*****************************************************************************************************************/	

If Not Exists (Select ProductId From CurrentStock Where ProductId=@DProductId)
  BEGIN
	Insert Into CurrentStock(ProductId, QtyLoose, Cost) Values(@DProductId, -@DQty, @DCost)	
  END
ELSE 
  begin
	If Exists (Select ProductId From CurrentStock Where ProductId=@DProductId and QtyLoose-@DQty<>0)
	  begin 
		Update CurrentStock Set QtyLoose=QtyLoose-(@DQty), 
		Cost=(@CurrentTTLCost - @DelTTLCost) / (QtyLoose-@DQty)
		Where ProductId=@DProductId
	  end
	else
	  begin
		Update CurrentStock Set QtyLoose=QtyLoose-(@DQty), 
		Cost=0
		Where ProductId=@DProductId
	  end
  end
if not Exists (Select ProductId From CurrentStockStore Where ProductId=@DProductId and StoreID=@DStoreId)
  BEGIN
		Insert Into CurrentStockStore(ProductId, StoreID, QtyLoose) Values(@DProductId, @DStoreId, -@DQty)		
  END
else
  BEGIN
		Update CurrentStockStore Set QtyLoose=QtyLoose-(@DQty)
		Where ProductId=@DProductId and StoreID=@DStoreId
  END

/***************************************************/
return
ERROR:
	raiserror (@ErrorMsg,16,1)
END