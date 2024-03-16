alter TRIGGER [td_ManufacturedProductsBody] ON dbo.ManufacturedProductsBody 
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
/********************************************/
DECLARE @MQtyLoose as numeric(11,4), @MRate as numeric(9,3)
/********************************************/
DECLARE @ProductId as varchar(5), @QtyLoose as numeric(9,3), @Cost as numeric(9,3)
/****************************************************/
DECLARE @DProductId as varchar(5), @DStoreID as tinyint, @DQty as numeric(9,3), @DCost as numeric(9,3),@DManufacturedID as int,
@CurrentTTLCost as numeric(18,3), @DelTTLCost as numeric(18,3)
/****************************************************/
Select @DProductId=ProductId, @DQty=QtyLoose, @DManufacturedID = ManufacturedID
From DELETED
Select @DStoreId=StoreId FROM ManufacturedProductsHeader WHERE ManufacturedID=(Select ManufacturedID FROM DELETED)
/***************************************************/
Select @CurrentTTLCost=round(Isnull((QtyLoose*Cost),0),3) From CurrentStock Where ProductId=@DProductId
/*****************************************************************************************************************/	
-- Declare the variables to store the values returned by FETCH.
set @DCost = 0
DECLARE Product_cursor CURSOR FOR
SELECT ProductID, QtyLoose 
FROM ProductProcessInfoHeader h 
inner join ProductProcessInfoBody b on h.ID = b.ID
where FinishedProductID = @DProductID
ORDER BY ProductID
OPEN Product_cursor

FETCH NEXT FROM Product_cursor
INTO @ProductID, @QtyLoose

-- Check @@FETCH_STATUS to see if there are any more rows to fetch.
WHILE @@FETCH_STATUS = 0
BEGIN
	select @DCost = @DCost + Isnull((@QtyLoose*Cost),0), @Cost=Cost From CurrentStock Where ProductId=@ProductId		
	if Exists (Select ProductId From ManufacturedProductsUsed Where ProductId=@ProductId and ManufacturedID=@DManufacturedID and QtyLoose = @QtyLoose*@DQty)
	  BEGIN
		delete from ManufacturedProductsUsed Where ProductId=@ProductId and ManufacturedID=@DManufacturedID
	  END
	else
	  BEGIN
		select @MRate = ((QtyLoose*Rate)-(@QtyLoose*@DQty*@Cost))/(QtyLoose-(@QtyLoose*@DQty)), @MQtyLoose = QtyLoose - (@QtyLoose*@DQty)
		From ManufacturedProductsUsed Where ProductId=@ProductId and ManufacturedID=@DManufacturedID
		Update ManufacturedProductsUsed Set QtyLoose = @MQtyLoose, 
		Rate = @MRate,
		Amount = @MQtyLoose * @MRate
		Where ProductId=@ProductId and ManufacturedID=@DManufacturedID
	  END
    -- This is executed as long as the previous fetch succeeds.
	FETCH NEXT FROM Product_cursor
	INTO @ProductID, @QtyLoose
END

CLOSE Product_cursor
DEALLOCATE Product_cursor
set @DelTTLCost = Round(@DQty * @DCost, 2)


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



