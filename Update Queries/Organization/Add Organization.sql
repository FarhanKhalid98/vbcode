ALTER TABLE OpeningStock ADD OrganizationID TinyInt Null 
GO
ALTER TABLE StockAdjustmentHeader ADD OrganizationID TinyInt Null 
GO
ALTER TABLE DisputeInvoiceBody ADD OrganizationID TinyInt Null 
GO
ALTER TABLE LiftInvoiceHeader ADD OrganizationID TinyInt Null 
GO