
-------------------- v1 Triger


Alter trigger tbi_SaleBody on dbo.SaleBody 
instead of insert as
begin
    IF EXISTS(Select 1 FROM INSERTED Where isProduct = 1)

        INSERT INTO SaleBody

                          (BillID, BillDate, ProductId, Code, PackingId, Multiplier, QtyPack, Qty, Bonus, Offer, SaleTaxPer, SaleTaxVal, Price, RetailPrice, IsWSSaleTax, IsRetailSaleTax, IsWSDiscb4ST, TokenVal, DiscPC, DiscPer, DiscVal, Cost, Amount, IsProduct )

        SELECT     BillID, BillDate, ProductId, Code, PackingId, Multiplier, QtyPack, Qty, Bonus, Offer, SaleTaxPer, SaleTaxVal, Price, RetailPrice, IsWSSaleTax, IsRetailSaleTax, IsWSDiscb4ST, TokenVal, DiscPC, DiscPer, DiscVal, 
						  (Case WHEN EXISTS(Select Cost FROM CurrentStock Where ProductId=INSERTED.ProductId) 
						  then (Select Cost FROM CurrentStock Where ProductId=INSERTED.ProductId)
							else 0 end), 
                          Amount, IsProduct 
                  FROM    INSERTED

    ELSE

        INSERT INTO SaleBody

                          (BillID, BillDate, ProductId, Code,PackingId, Multiplier, QtyPack, Qty, Bonus, Offer, SaleTaxPer, SaleTaxVal, Price, RetailPrice, IsWSSaleTax, IsRetailSaleTax, IsWSDiscb4ST, TokenVal, DiscPC, DiscPer, DiscVal, Cost, Amount, IsProduct)

        SELECT      BillID, BillDate, ProductId, Code, PackingId, Multiplier, QtyPack, Qty, Bonus, Offer, SaleTaxPer, SaleTaxVal, Price, RetailPrice, IsWSSaleTax, IsRetailSaleTax, IsWSDiscb4ST, TokenVal, DiscPC, DiscPer, DiscVal,
						(0), 
						Amount, IsProduct 

                  FROM    INSERTED

end



