DECLARE @ProductId as varchar(5), @PurPrice as numeric(9,3)
DECLARE Product_cursor CURSOR FOR
SELECT ProductID, PurPrice 
FROM Products
ORDER BY ProductID
OPEN Product_cursor

FETCH NEXT FROM Product_cursor
INTO @ProductID, @PurPrice

-- Check @@FETCH_STATUS to see if there are any more rows to fetch.
WHILE @@FETCH_STATUS = 0
BEGIN		

	update currentstock set cost = @PurPrice where ProductID = @ProductID
   -- This is executed as long as the previous fetch succeeds.
	FETCH NEXT FROM Product_cursor
	INTO @ProductID, @PurPrice
END

CLOSE Product_cursor
DEALLOCATE Product_cursor
