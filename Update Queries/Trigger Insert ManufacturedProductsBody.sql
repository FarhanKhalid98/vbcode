create TRIGGER [ti_ManufacturedProductsBody] ON dbo.ManufacturedProductsBody 
FOR INSERT
AS
BEGIN

DECLARE @ErrorMsg as varchar(100)
If (@@ROWCOUNT>1)
BEGIN
	set @ErrorMsg= 'Multiple rows cannot be inserted.'
	goto error
END
/********************************************/
DECLARE @MQtyLoose as numeric(11,4), @MRate as numeric(9,3)
/********************************************/
DECLARE @ProductId as varchar(5), @QtyLoose as numeric(11,4), @Cost as numeric(9,3)
/***** Declaring variables to be inserted *******/
DECLARE @IProductId as varchar(5), @IStoreID as tinyint, @IQty as numeric(9,4), @ICost as numeric(9,3), @IManufacturedID as int,
@CurrentTTLCost as numeric(18,6), @NewTTLCost as numeric(18,6)
/****************************************************/
Select @IProductId=ProductId, @IQty = QtyLoose, @IManufacturedID = ManufacturedID
From INSERTED 
Select @IStoreId=StoreId FROM ManufacturedProductsHeader WHERE ManufacturedID=(Select ManufacturedID FROM Inserted)
/***************************************************/
Select @CurrentTTLCost=Isnull((QtyLoose*Cost),0) From CurrentStock Where ProductId=@IProductId
/***************************************************************************/
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
If Not Exists (Select ProductId From CurrentStock Where ProductId=@IProductId)
  BEGIN
	Insert Into CurrentStock(ProductId, QtyLoose, Cost) Values(@IProductId, @IQty, @ICost)	
  END
ELSE 
  begin
	If Exists (Select ProductId From CurrentStock Where ProductId=@IProductId and QtyLoose+@IQty<>0)
	  begin 
		Update CurrentStock Set QtyLoose=QtyLoose+(@IQty), 
		Cost=(@CurrentTTLCost + @NewTTLCost) / (QtyLoose+@IQty)
		Where ProductId=@IProductId
	  end
	else
	  begin
		Update CurrentStock Set QtyLoose=QtyLoose+(@IQty), 
		Cost=0
		Where ProductId=@IProductId
	  end
  end
if not Exists (Select ProductId From CurrentStockStore Where ProductId=@IProductId and StoreID=@IStoreId)
  BEGIN
		Insert Into CurrentStockStore(ProductId, StoreID, QtyLoose) Values(@IProductId, @IStoreId, @IQty)		
  END
else
  BEGIN
		Update CurrentStockStore Set QtyLoose=QtyLoose+(@IQty)
		Where ProductId=@IProductId and StoreID=@IStoreId
  END

/***************************************************/
return
ERROR:
	raiserror (@ErrorMsg,16,1)
END