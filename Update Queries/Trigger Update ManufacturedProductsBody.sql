CREATE TRIGGER [tu_ManufacturedProductsBody] ON dbo.ManufacturedProductsBody 
FOR UPDATE
AS
BEGIN
/***** Declaring variables to be inserted *******/
DECLARE @ErrorMsg as varchar(100)

If (@@ROWCOUNT>1)
BEGIN
	set @ErrorMsg= 'Multiple rows cannot be updated.'
	goto error
END
/********************************************/
DECLARE @MQtyLoose as numeric(11,4), @MRate as numeric(9,3)
/************************************/
DECLARE @ProductId as varchar(5), @QtyLoose as numeric(11,4), @Cost as numeric(9,3)
/*************CURRENT STOCK QTY*******************************/
DECLARE @CSQty as numeric(9,3),@CSCost as numeric(9,3),  @CSTTLCost as numeric(18,6)
/*************CURRENT STOCK STORE QTY*******************************/
DECLARE @CSSQty as numeric(9,3)
/***** Declaring variables to be inserted *******/
DECLARE @IProductId as varchar(5), @IStoreID as tinyint, @IQty as numeric(9,3), @ICost as numeric(9,3), @IManufacturedID as int,
@NewTTLCost as numeric(18,6)
/***** Declaring variables to be Deleted *******/
DECLARE @DProductId as varchar(5), @DStoreID as tinyint, @DQty as numeric(9,3), @DCost as numeric(9,3), @DManufacturedID as int,
@DelTTLCost as numeric(18,6)
/**************************************************************************/
Select @IProductId=ProductId, @IQty = QtyLoose, @IManufacturedID = ManufacturedID
From INSERTED 
Select @IStoreId=StoreId FROM ManufacturedProductsHeader WHERE ManufacturedID=(Select ManufacturedID FROM Inserted)
/****************************************************/
Select @DProductId=ProductId, @DQty=QtyLoose, @DManufacturedID = ManufacturedID
From DELETED
Select @DStoreId=StoreId FROM ManufacturedProductsHeader WHERE ManufacturedID=(Select ManufacturedID FROM DELETED)
/***************************************************/
Select @CSSQty=QtyLoose From CurrentStockStore Where ProductId=@DProductId and StoreID=@DStoreID
/***************************************************/
Select @CSQty=QtyLoose, @CSCost=Cost, @CSTTLCost=QtyLoose*Cost From CurrentStock Where ProductId=@IProductId
/***************************************************/
--deleted
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
/*********************************************/
-- inserted
-- Declare the variables to store the values returned by FETCH.
set @ICost = 0
DECLARE Product_cursor CURSOR FOR
SELECT ProductID, QtyLoose 
FROM ProductProcessInfoHeader h 
inner join ProductProcessInfoBody b on h.ID = b.ID
where FinishedProductID = @IProductID
ORDER BY ProductID
OPEN Product_cursor

FETCH NEXT FROM Product_cursor
INTO @ProductID, @QtyLoose

-- Check @@FETCH_STATUS to see if there are any more rows to fetch.
WHILE @@FETCH_STATUS = 0
BEGIN		
	select @ICost = @ICost + Isnull((@QtyLoose*Cost),0), @Cost=Cost From CurrentStock Where ProductId=@ProductId	
	-- update ManufacturedProductsUsed
	if not Exists (Select ProductId From ManufacturedProductsUsed Where ProductId=@ProductId and ManufacturedID=@IManufacturedID)
	  BEGIN
		Insert Into ManufacturedProductsUsed(ManufacturedID, ProductId, QtyLoose, Rate, Amount) Values(@IManufacturedID, @ProductId, @QtyLoose*@IQty, @Cost, @QtyLoose*@IQty*@Cost)		
	  END
	else
	  BEGIN
		select @MRate = ((QtyLoose*Rate)+(@QtyLoose*@IQty*@Cost))/(QtyLoose+(@QtyLoose*@IQty)), @MQtyLoose = QtyLoose + (@QtyLoose*@IQty)
		From ManufacturedProductsUsed Where ProductId=@ProductId and ManufacturedID=@IManufacturedID
		Update ManufacturedProductsUsed Set QtyLoose = @MQtyLoose, 
		Rate = @MRate,
		Amount = @MQtyLoose * @MRate
		Where ProductId=@ProductId and ManufacturedID=@IManufacturedID
	  END
   -- This is executed as long as the previous fetch succeeds.
	FETCH NEXT FROM Product_cursor
	INTO @ProductID, @QtyLoose
END

CLOSE Product_cursor
DEALLOCATE Product_cursor
set @NewTTLCost = Round(@IQty * @ICost, 2)
/****************************************************/
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
	raiserror (@ErrorMsg,16,1)
END



