CREATE TRIGGER [td_SaleBody] ON dbo.SaleBody 
FOR DELETE
AS
BEGIN

DECLARE @ErrorMsg as varchar(100)
If (@@ROWCOUNT>1)
BEGIN
	set @ErrorMsg= 'Multiple rows cannot be inserted.'  +  cast(@@ROWCOUNT as varchar(3))
	goto error
END
/***** Declaring variables to be inserted *******/
DECLARE @DProductId as varchar(5), @DStoreID as tinyint, @DQty as numeric(9,2), @DCost as numeric(9,3),
@CurrentTTLCost as numeric(18,6), @DelTTLCost as numeric(18,6)
/****************************************************/
Select @DProductId=ProductId, @DStoreId=StoreId, @DQty=Qty, @DCost=Cost, @DelTTLCost = Round((@DQty)*@DCost, 2)
From DELETED d inner join SaleHeader h on d.billid = h.billid and d.billdate = h.billdate
/***************************************************/
Select @CurrentTTLCost=Isnull((QtyLoose*Cost),0) From CurrentStock Where ProductId=@DProductId
/*****************************************************************************************************************/
If Not Exists (Select ProductId From CurrentStock Where ProductId=@DProductId)
  BEGIN
	Insert Into CurrentStock(ProductId, QtyLoose, Cost) Values(@DProductId, @DQty, @DCost)	
  END
ELSE 
  begin
	If Exists (Select ProductId From CurrentStock Where ProductId=@DProductId and QtyLoose+@DQty<>0)
	  begin 
		Update CurrentStock Set QtyLoose=QtyLoose+(@DQty), 
		Cost=(@CurrentTTLCost + @DelTTLCost) / (QtyLoose+@DQty)
		Where ProductId=@DProductId
	  end
	else
	  begin
		Update CurrentStock Set QtyLoose=QtyLoose+(@DQty), 
		Cost=0
		Where ProductId=@DProductId
	  end
  end
if not Exists (Select ProductId From CurrentStockStore Where ProductId=@DProductId and StoreID=@DStoreId)
  BEGIN
		Insert Into CurrentStockStore(ProductId, StoreID, QtyLoose) Values(@DProductId, @DStoreId, @DQty)		
  END
else
  BEGIN
		Update CurrentStockStore Set QtyLoose=QtyLoose+(@DQty)
		Where ProductId=@DProductId and StoreID=@DStoreId
  END
/***************************************************/
return
ERROR:
	raiserror (@ErrorMsg,16,1)
END