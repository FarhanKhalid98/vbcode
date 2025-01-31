if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ProdRptProductStockValue]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ProdRptProductStockValue]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE	PROCEDURE ProdRptProductStockValue
	@FromDate smalldatetime = '1/1/1980', 
	@ToDate   smalldatetime = '1/1/2020',
	@ProductID  VARCHAR (5)= null,
	@GroupID  VARCHAR (3)= null,
	@CompanyID  VARCHAR (3)= null,
	@SubGroupID  VARCHAR (3)= null
AS
select d.ProductID, ProductName, p.GroupID, GroupName, p.CompanyID, CompanyName, sum(op) as OP, sum(p) as P, sum(pr) as PR, sum(s) as S, sum(sr) as SR, sum(AdjIn) as AdjIn, sum(AdjOut) as AdjOut, Sum(LI) as LI, Sum(DOvr) as DOvr, Sum(DUnd) as DUnd,
       Case When Sum(OP) = 0 Then 0 Else sum(OP_ProdNetAmount) End OP_ProdNetAmount, Sum(P_ProdNetAmount) P_ProdNetAmount, Sum(PR_ProdNetAmount) PR_ProdNetAmount, Sum(S_ProdNetAmount) S_ProdNetAmount, Sum(SR_ProdNetAmount) SR_ProdNetAmount, Sum(AdjIn_ProdNetAmount) AdjIn_ProdNetAmount, Sum(AdjOut_ProdNetAmount) AdjOut_ProdNetAmount, Sum(LI_ProdNetAmount) LI_ProdNetAmount, Sum(DOvr_ProdNetAmount) DOvr_ProdNetAmount, Sum(DUnd_ProdNetAmount) DUnd_ProdNetAmount,

       Case When Sum(OP) = 0 Then 0 Else sum(OP_ProdNetAmount) / Sum(OP) end as OP_Price, 

       case when Sum(p) = 0 then 0 Else Sum(P_ProdNetAmount) /  Sum(P) end as P_Price, 

       case when Sum(pr) = 0  then 0 else Sum(PR_ProdNetAmount) /  Sum(Pr) end as PR_Price,

       Case When Sum(S)= 0 then 0 Else Sum(S_ProdNetAmount) /  Sum(S) End  as S_Price, 

       case when Sum(sr) = 0 then 0 else Sum(SR_ProdNetAmount) /  Sum(Sr) end as SR_Price, 

       case when Sum(Adjin) = 0 then 0 else Sum(AdjIn_ProdNetAmount) / Sum(AdjIn) end as AdjIn_Price, 

       case When Sum(AdjOut) = 0 then 0 else Sum(AdjOut_ProdNetAmount) / Sum(AdjOut) End as AdjOut_Price, 

       case when Sum(LI) = 0 then 0 else Sum(LI_ProdNetAmount) / Sum(LI) End as LI_Price, Sum(DOvr_ProdNetAmount) as DOvr_Price, Sum(DUnd_ProdNetAmount) as DUnd_Price
from
(

select ProductID, sum(op) + sum(p) - sum(pr) - sum(AdjOut) - sum(s) + sum(sr) + sum(AdjIn) + Sum(DOvr) - Sum(DUnd) as OP, 
	          Sum(OP_ProdNetAmount) + Sum(P_ProdNetAmount) -  Sum(PR_ProdNetAmount) - Sum(AdjOut_ProdNetAmount) - Sum(S_ProdNetAmount) + Sum(SR_ProdNetAmount) + Sum(AdjIn_ProdNetAmount) + Sum(DOvr_ProdNetAmount) -Sum(DUnd_ProdNetAmount) as OP_ProdNetAmount,
	0 as P, 0 as PR, 0 as S, 0 as SR, 0 as AdjIn, 0 as AdjOut, 0 as LI, 0 as DOvr, 0 as DUnd,
        0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, 0 as SR_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from 
(
select ProductID, qtyLoose as OP, Amount as OP_ProdNetAmount, 0 as P, 0 as PR, 0 as S, 0 as SR, 0 as AdjIn, 0 as AdjOut,0 as LI,0 as DOvr, 0 as DUnd,
       0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, 0 as SR_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount  
from OpeningStock

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, isnull(multiplier,0)*isnull(qtypack,0) + qtyLoose + isnull(bonus,0) as P, 0 as PR, 0 as S, 0 as SR, 0 as AdjIn, 0 as AdjOut, 0 as LI, 0 as DOvr, 0 as DUnd,
       (Amount-(Amount* (isnull(BillDiscPer,0)/100))) +  case When totalamount = 0 then 0 else (Amount* isnull(OtherCharges,0) /TotalAmount) + (Amount* isnull(TotalExpense,0) /TotalAmount) end as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, 0 as SR_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from PurchaseBody b
Inner Join PurchaseHeader h on h.purid = b.purid and h.purchasedate = b.purchaseDate
where h.PurchaseDate < @FromDate

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, isnull(multiplier,0)*isnull(qtypack,0) + qtyLoose + isnull(bonus,0) as PR, 0 as S, 0 as SR, 0 as AdjIn, 0 as AdjOut, 0 as LI, 0 as DOvr, 0 as DUnd,
       0 as P_ProdNetAmount, (Amount-(Amount* (isnull(BillDiscPer,0)/100))) +  case When totalamount = 0 then 0 else (Amount* isnull(0,0) /TotalAmount) + (Amount* isnull(TotalExpense,0) /TotalAmount) end as Pr_ProdNetAmount, 0 as S_ProdNetAmount, 0 as SR_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from PurchaseReturnBody B
Inner Join PurchaseReturnHeader h on h.returnid = b.returnid and h.returndate = b.ReturnDate
where h.ReturnDate < @FromDate

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, 0 as PR, Qty as S, 0 as SR, 0 as AdjIn, 0 as AdjOut, 0 as LI, 0 as DOvr, 0 as DUnd,
       0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, Qty * Cost as S_ProdNetAmount, 0 as SR_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from SaleBody b
inner join SaleHeader H on h.billid = b.billid and h.billdate = b.billdate 
where h.BillDate < @FromDate and isproduct = 1

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, 0 as PR, QtyLoose as S, 0 as SR, 0 as AdjIn, 0 as AdjOut, 0 as LI, 0 as DOvr, 0 as DUnd,
       0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, QtyLoose * Rate as S_ProdNetAmount, 0 as SR_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from SaleUnionUsed b
inner join SaleHeader H on h.billid = b.billid and h.billdate = b.billdate 
where h.BillDate < @FromDate

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, 0 as PR, 0 as S, Qty as SR, 0 as AdjIn, 0 as AdjOut, 0 as LI, 0 as DOvr, 0 as DUnd,
       0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, Qty * Cost as Sr_ProdNetAmount,  0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from SaleReturnBody B
inner join SaleReturnHeader H on h.Returnid = b.Returnid and h.returndate = b.returndate 
where h.ReturnDate < @FromDate and isproduct = 1

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, 0 as PR, 0 as S, QtyLoose as SR, 0 as AdjIn, 0 as AdjOut, 0 as LI, 0 as DOvr, 0 as DUnd,
       0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, QtyLoose * Rate as Sr_ProdNetAmount,  0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from SaleReturnUnionUsed B
inner join SaleReturnHeader H on h.Returnid = b.Returnid and h.returndate = b.returndate 
where h.ReturnDate < @FromDate

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, 0 as PR, 0 as S, 0 as SR,
isnull(multiplier,0)* isnull(OQtyPack,0) + OQtyLoose as AdjIn,
isnull(multiplier,0)* isnull(UQtyPack,0) + UQtyLoose as AdjOut, 0 as LI, 0 as DOvr, 0 as DUnd,
0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, 0 as Sr_ProdNetAmount, 
isnull(multiplier,0)* isnull(OQtyPack,0) + OQtyLoose * Cost  as AdjIn_ProdNetAmount, 
isnull(multiplier,0)* isnull(UQtyPack,0) + UQtyLoose  * Cost as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from StockAdjustmentHeader h inner join
StockAdjustmentBody b on h.AdjustmentID = b.AdjustmentID
where AdjustmentDate < @FromDate

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, 0 as PR, 0 as S, 0 as SR, 0 as AdjIn, 0 as AdjOut,  isnull(multiplier,0)*isnull(qtypack,0) + qtyLoose as LI, 0 as DOvr, 0 as DUnd,
       0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, 0 as Sr_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount,  amount as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from LiftInvoiceBody
where LiftDate < @FromDate

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, 0 as PR, 0 as S, 0 as SR, 0 as AdjIn, 0 as AdjOut,  0 as LI, OverQty as DOvr, UnderQty as DUnd,
       0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, 0 as Sr_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount
from DisputeInvoiceBody
where DisputeDate < @FromDate

) d
group by ProductID

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, isnull(multiplier,0)*isnull(qtypack,0) + qtyLoose + isnull(bonus,0) as P, 0 as PR, 0 as S, 0 as SR, 0 as AdjIn, 0 as AdjOut,0 as LI, 0 as DOvr, 0 as DUnd,
       (Amount-(Amount* (isnull(BillDiscPer,0)/100))) +  case When totalamount = 0 then 0 else (Amount* isnull(OtherCharges,0) /TotalAmount) + (Amount* isnull(TotalExpense,0) /TotalAmount) end as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, 0 as SR_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from PurchaseBody b
Inner Join PurchaseHeader h on h.purid = b.purid and h.purchasedate = b.purchaseDate
where h.PurchaseDate between @FromDate and @ToDate

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P,  isnull(multiplier,0)*isnull(qtypack,0) + qtyLoose + isnull(bonus,0)  as PR, 0 as S, 0 as SR, 0 as AdjIn, 0 as AdjOut,0 as LI , 0 as DOvr, 0 as DUnd,
       0 as P_ProdNetAmount, (Amount-(Amount* (isnull(BillDiscPer,0)/100))) +  case When totalamount = 0 then 0 else (Amount* isnull(0,0) /TotalAmount)+ (Amount* isnull(TotalExpense,0) /TotalAmount) end as Pr_ProdNetAmount, 0 as S_ProdNetAmount, 0 as SR_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from PurchaseReturnBody B
Inner Join PurchaseReturnHeader h on h.returnid = b.returnid and h.returndate = b.ReturnDate
where h.ReturnDate between @FromDate and @ToDate

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, 0 as PR, Qty as S, 0 as SR, 0 as AdjIn, 0 as AdjOut,0 as LI, 0 as DOvr, 0 as DUnd,
       0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, Qty * Cost  as S_ProdNetAmount, 0 as SR_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from SaleBody b
inner join SaleHeader H on h.billid = b.billid and h.billdate = b.billdate 
where h.BillDate between @FromDate and @ToDate and isProduct = 1

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, 0 as PR, QtyLoose as S, 0 as SR, 0 as AdjIn, 0 as AdjOut,0 as LI, 0 as DOvr, 0 as DUnd,
       0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, QtyLoose * Rate  as S_ProdNetAmount, 0 as SR_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from SaleUnionUsed b
inner join SaleHeader H on h.billid = b.billid and h.billdate = b.billdate 
where h.BillDate between @FromDate and @ToDate

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, 0 as PR, 0 as S, Qty as SR, 0 as AdjIn, 0 as AdjOut,0 as LI, 0 as DOvr, 0 as DUnd,
       0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, Qty * Cost as Sr_ProdNetAmount,  0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from SaleReturnBody B
inner join SaleReturnHeader H on h.Returnid = b.Returnid and h.returndate = b.returndate 
where h.ReturnDate between @FromDate and @ToDate and isProduct = 1

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, 0 as PR, 0 as S, QtyLoose as SR, 0 as AdjIn, 0 as AdjOut,0 as LI, 0 as DOvr, 0 as DUnd,
       0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, QtyLoose * Rate  as Sr_ProdNetAmount,  0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from SaleReturnUnionUsed B
inner join SaleReturnHeader H on h.Returnid = b.Returnid and h.returndate = b.returndate 
where h.ReturnDate between @FromDate and @ToDate

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, 0 as PR, 0 as S, 0 as SR,
isnull(multiplier,0)* isnull(OQtyPack,0) + OQtyLoose as AdjIn, 
isnull(multiplier,0)* isnull(UQtyPack,0) + UQtyLoose as AdjOut ,0 as LI, 0 as DOvr, 0 as DUnd,
0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, 0 as Sr_ProdNetAmount, 
isnull(multiplier,0)* isnull(OQtyPack,0) + OQtyLoose * Cost  as AdjIn_ProdNetAmount, 
isnull(multiplier,0)* isnull(UQtyPack,0) + UQtyLoose  * Cost as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from StockAdjustmentHeader h inner join
StockAdjustmentBody b on h.AdjustmentID = b.AdjustmentID
where AdjustmentDate between @FromDate and @ToDate

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount,0 as P, 0 as PR, 0 as S, 0 as SR, 0 as AdjIn, 0 as AdjOut,  isnull(multiplier,0)*isnull(qtypack,0) + qtyLoose as LI, 0 as DOvr, 0 as DUnd,
       0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, 0 as Sr_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount,  amount as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount 
from LiftInvoiceBody
where LiftDate between @FromDate and @ToDate

Union All

select ProductID, 0 as OP, 0 as OP_ProdNetAmount, 0 as P, 0 as PR, 0 as S, 0 as SR, 0 as AdjIn, 0 as AdjOut,  0 as LI, OverQty as DOvr, UnderQty as DUnd,
       0 as P_ProdNetAmount, 0 as PR_ProdNetAmount, 0 as S_ProdNetAmount, 0 as Sr_ProdNetAmount, 0 as AdjIn_ProdNetAmount, 0 as AdjOut_ProdNetAmount, 0 as LI_ProdNetAmount, 0 as DOvr_ProdNetAmount, 0 as DUnd_ProdNetAmount
from DisputeInvoiceBody
where DisputeDate Between @FromDate and @ToDate

)d inner join Products p on d.ProductID = p.ProductID
inner join Groups g on g.GroupID = p.GroupID
Left Outer join Companies c on c.CompanyID = p.CompanyID
where (@ProductID is null or d.ProductID = @ProductID)
and (@GroupID is null or p.GroupID = @GroupID)
and (@CompanyID is null or p.CompanyID = @CompanyID)
and (@SubGroupID is null or p.SubGroupID = @SubGroupID)
group by d.ProductID, ProductName, p.GroupID, GroupName, p.CompanyID, CompanyName

order by d.productid


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

