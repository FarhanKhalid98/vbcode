ALTER TABLE Registry ADD SetCurrentStock bit NULL
go
update Registry set SetCurrentStock =1
go
ALTER TABLE Registry ALTER COLUMN SetCurrentStock bit not null
go

ALTER TABLE Registry ADD CostVisible bit NULL
go
update Registry set CostVisible =1
go
ALTER TABLE Registry ALTER COLUMN CostVisible bit not null
go