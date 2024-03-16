Attribute VB_Name = "AdminClosing"
Option Explicit
Dim sSql As String
Public vTotalSale, vPettyCash, vRecoveryCustomer, vBankCardSale, vCreditSale, vDiscount, vServiceCharges As Double
Public vSTax, vSaleReturn, vPayments, vCashReceived, vCreditSaleReturnPaid, vBankReceived, vBankPayments, vCashAvailable, vExcessShort As Double

Public Function CalculateAmount(vUserID As Byte, vEntryDate As Date)
On Error GoTo ErrorHandler
   ' Step 1 - Total Sale
   sSql = " Select isnull(round(Sum((isnull(multiplier,0)* isnull(QtyPack,0) + Qty )*((Price/isnull(multiplier,1))+isnull(sc,0))),0),0) as TotalSale" & vbCrLf _
      + " from SaleHeader h inner join SaleBody b on H.SID = B.SID and h.billdate = b.billdate" & vbCrLf _
      + " where 1=1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.BillDate = '" & vEntryDate & "'"
   vTotalSale = CN.Execute(sSql).Fields(0).Value

   sSql = " Select isnull(Sum(totalamount),0) as TotalSale" & vbCrLf _
      + " from CustomOrderHeader " & vbCrLf _
      + " where 1=1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and OrderDate = '" & vEntryDate & "'"
   vTotalSale = Val(vTotalSale) + CN.Execute(sSql).Fields(0).Value
   
'   sSql = " Select isnull(floor(Sum(Qty*Price)),0) as TotalSale" & vbCrLf _
      + " from ServiceHeader h inner join ServiceBody b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
      + " where 1=1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.BillDate = '" & VEntryDate & "'"
'   vTotalSale = Val(vTotalSale) + CN.Execute(sSql).Fields(0).Value
   
   ' Step 2 - Total PettyCash
   sSql = " Select isnull(sum(Amount),0)amount from PettyCashHeader where 1=1 " & IIf(vUserID = 0, "", " and ToUserNo = " & vUserID) & " and EntryDate = '" & vEntryDate & "'"
   vPettyCash = CN.Execute(sSql).Fields(0).Value

   ' Step 3 - Total Customer Recovery
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " FROM RecoveryHeader h INNER JOIN RecoveryCustomer b ON h.RecoveryId = B.RecoveryId " & vbCrLf _
      + " where BankMachineiD is null  " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.RecoveryDate = '" & vEntryDate & "'"
   vRecoveryCustomer = CN.Execute(sSql).Fields(0).Value

   sSql = " Select isnull(sum(Payment),0) as Amount " & vbCrLf _
      + " FROM CustomOrderDelivery" & vbCrLf _
      + " where Cash=1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and DeliveryDate = '" & vEntryDate & "'"
   vRecoveryCustomer = Val(vRecoveryCustomer) + CN.Execute(sSql).Fields(0).Value

   ' Step 4 - Bank Card Sale
   sSql = " Select isnull(Sum(Amount + isnull(ServiceCharges,0) + isnull(STax,0) - isnull(CashReceived,0) - isnull(BillDisc,0)),0)  as TotalBankSale" & vbCrLf _
      + " from SaleHeader h inner join (select SID, sum(Amount) as Amount From SaleBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where BankCard = 1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.BillDate = '" & vEntryDate & "'"
   vBankCardSale = CN.Execute(sSql).Fields(0).Value
   
    sSql = " Select isnull(Sum(BankAmount),0)  as TotalBankSale" & vbCrLf _
      + " from SaleHeader h inner join (select SID, sum(Amount) as Amount From SaleBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where credit = 1  and isnull(bankamount,0) > 0 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.BillDate = '" & vEntryDate & "'"
   vBankCardSale = Val(vBankCardSale) + CN.Execute(sSql).Fields(0).Value


   sSql = " Select isnull(Sum(Amount - isnull(BillDisc,0) + isnull(ServiceCharges,0) + isnull(STax,0) ),0)  as TotalBankSale" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) as Amount From SaleReturnBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where BankCard = 1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.ReturnDate = '" & vEntryDate & "'"
   vBankCardSale = Val(vBankCardSale) - CN.Execute(sSql).Fields(0).Value
   
    sSql = " Select isnull(Sum(BankAmount),0)  as TotalBankSale" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) as Amount From SaleReturnBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where credit = 1  and isnull(bankamount,0) > 0  " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.ReturnDate = '" & vEntryDate & "'"
   vBankCardSale = Val(vBankCardSale) - CN.Execute(sSql).Fields(0).Value

   sSql = " Select isnull(Sum(Amount - isnull(BillDisc,0)),0)  as TotalBankSale" & vbCrLf _
      + " from ServiceHeader h inner join (select BillID, BillDate, sum(Amount) as Amount From ServiceBody Group By BillId, BillDate) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
      + " where BankCard = 1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.BillDate = '" & vEntryDate & "'"
   vBankCardSale = Val(vBankCardSale) - CN.Execute(sSql).Fields(0).Value

   ' Step 5 - Credit Sale
   sSql = " Select isnull(round(sum(Amount - isnull(CashReceived,0) - isnull(BillDisc,0) + isnull(ServiceCharges,0) + isnull(OtherCharges,0) + isnull(STax,0) ),0),0) as CreditSale" & vbCrLf _
      + " from SaleHeader h inner join (select SID, sum(Amount) as Amount From SaleBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where Credit = 1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.BillDate = '" & vEntryDate & "'"
   vCreditSale = CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(Sum(BankAmount),0)  as TotalBankSale" & vbCrLf _
      + " from SaleHeader h inner join (select SID, sum(Amount) as Amount From SaleBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where credit = 1  and isnull(bankamount,0) > 0 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.BillDate = '" & vEntryDate & "'"
   vCreditSale = Val(vCreditSale) - CN.Execute(sSql).Fields(0).Value
   
    sSql = " Select isnull(round(sum(Amount - isnull(CashPaid,0) - isnull(BillDisc,0) + isnull(ServiceCharges,0) + isnull(OtherCharges,0) + isnull(STax,0) ),0),0) as CreditSale" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) amount From SaleReturnBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where Credit = 1  " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.ReturnDate = '" & vEntryDate & "'"
   vCreditSale = Val(vCreditSale) - CN.Execute(sSql).Fields(0).Value
      
   sSql = " Select isnull(Sum(BankAmount),0)  as TotalBankSale" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) as Amount From SaleReturnBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where credit = 1  and isnull(bankamount,0) > 0  " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.ReturnDate = '" & vEntryDate & "'"
  vCreditSale = Val(vCreditSale) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(TotalAmount-isnull(Advance,0)),0) as CreditSale" & vbCrLf _
      + " from CustomOrderHeader " & vbCrLf _
      + " where Cash = 1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and OrderDate = '" & vEntryDate & "'"
   vCreditSale = Val(vCreditSale) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(TotalAmount),0) as CreditSale" & vbCrLf _
      + " from CustomOrderHeader " & vbCrLf _
      + " where Credit = 1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and OrderDate = '" & vEntryDate & "'"
   vCreditSale = Val(vCreditSale) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(round(sum(Amount - isnull(CashReceived,0) - isnull(BillDisc,0) ),0),0) as CreditSale" & vbCrLf _
      + " from ServiceHeader h inner join (select BillID, BillDate, sum(Amount) as Amount From ServiceBody Group By BillId, BillDate) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
      + " where Credit = 1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.BillDate = '" & vEntryDate & "'"
   vCreditSale = Val(vCreditSale) + CN.Execute(sSql).Fields(0).Value

   ' Step 6 - Discount
   sSql = " Select isnull(floor(isnull(sum(BillDisc),0) + isnull(sum(discval),0)),0) as Discount" & vbCrLf _
      + " from SaleHeader h inner join (select SID, sum(discval)discval From SaleBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where 1 = 1  " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.BillDate = '" & vEntryDate & "'"
   vDiscount = CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(floor(isnull(sum(BillDisc),0) + isnull(sum(discval),0)),0) as Discount" & vbCrLf _
      + " from ServiceHeader h inner join (select BillId, BillDate, sum(discval)discval From ServiceBody Group By BillId, BillDate) b on h.BillID = b.BillID and h.BillDate = b.BillDate " & vbCrLf _
      + " where 1 = 1  " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.BillDate = '" & vEntryDate & "'"
   vDiscount = Val(vDiscount) + CN.Execute(sSql).Fields(0).Value
   
   ' Step 7 - Service Charges
   sSql = " Select isnull(sum(isnull(ServiceCharges,0)+ isnull(othercharges,0)),0)  as ServiceCharges" & vbCrLf _
      + " from SaleHeader h inner join (select SID, sum(discval)discval From SaleBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where 1 = 1  " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.BillDate = '" & vEntryDate & "'"
   vServiceCharges = CN.Execute(sSql).Fields(0).Value
   
   ' Step 8 - Sales Tax
   sSql = " Select isnull(sum(STax),0) as STax" & vbCrLf _
      + " from SaleHeader h inner join (select SID, sum(discval)discval From SaleBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where 1 = 1  " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.BillDate = '" & vEntryDate & "'"
   vSTax = CN.Execute(sSql).Fields(0).Value
   
   ' Step 9 - Sale Return
   
   
   sSql = " Select isnull(round(sum(Amount - isnull(BillDisc,0) + isnull(ServiceCharges,0) + isnull(STax,0) ),0),0) as SaleReturn" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) amount From SaleReturnBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where (Cash = 1 or BankCard = 1) and 1=1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.ReturnDate = '" & vEntryDate & "'"
   vSaleReturn = CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(round(sum(Amount - isnull(CashPaid,0) - isnull(BillDisc,0) + isnull(ServiceCharges,0) + isnull(STax,0) ),0),0) as SaleReturn" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) amount From SaleReturnBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where Credit = 1 and 1=1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.ReturnDate = '" & vEntryDate & "'"
   vSaleReturn = Val(vSaleReturn) + CN.Execute(sSql).Fields(0).Value
   
   
   ' Step 10 - Total Payments
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " FROM DebitVouchers h INNER JOIN DebitVouchersBody b ON h.VoucherNo = B.VoucherNo and h.Storeid = b.Storeid" & vbCrLf _
      + " where BankID is null " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.VoucherDate = '" & vEntryDate & "'"
   vPayments = CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(PaidAmount),0) as Amount " & vbCrLf _
      + " FROM PurchaseHeader " & vbCrLf _
      + " where 1=1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and PurchaseDate = '" & vEntryDate & "'"
   vPayments = Val(vPayments) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " from PaymentHeader h inner join PaymentVender v on h.PaymentID = v.PaymentID " & vbCrLf _
      + " where BankID is null  " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and PaymentDate = '" & vEntryDate & "'"
   vPayments = Val(vPayments) + CN.Execute(sSql).Fields(0).Value
   
'   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
'      + " FROM RecoveryHeader h INNER JOIN RecoveryCustomer b ON h.RecoveryId = B.RecoveryId " & vbCrLf _
'      + " where 1=1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.RecoveryDate = '" & vEntryDate & "'"
'   vPayments = Val(vPayments) + cn.Execute(sSql).Fields(0).Value

   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " from AdvanceVouchers h inner join AdvanceVouchersBody b on h.VoucherNo = b.VoucherNo" & vbCrLf _
      + " where 1=1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.VoucherDate = '" & vEntryDate & "'"
   vPayments = Val(vPayments) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " FROM DebitVouchers h INNER JOIN DebitVouchersBody b ON h.VoucherNo = B.VoucherNo and h.Storeid = b.Storeid" & vbCrLf _
      + " where 1=1 and BankId is Not null " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.VoucherDate = '" & vEntryDate & "'"
   vBankPayments = CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " from PaymentHeader h inner join PaymentVender v on h.PaymentID = v.PaymentID " & vbCrLf _
      + " where BankID is Not  null  " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and PaymentDate = '" & vEntryDate & "'"
   vBankPayments = Val(vBankPayments) + CN.Execute(sSql).Fields(0).Value
   
   ' Step 11 - Total Received Payments
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " FROM CreditVouchers h INNER JOIN CreditVouchersBody b ON h.VoucherNo = B.VoucherNo and h.Storeid = b.Storeid" & vbCrLf _
      + " where 1=1 and BankId is null " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.VoucherDate = '" & vEntryDate & "'"
   vCashReceived = CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(CashReceived),0) as CreditSale" & vbCrLf _
      + " from SaleOrderHeader h " & vbCrLf _
      + " where Credit = 1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.OrderDate = '" & vEntryDate & "'"
   vCashReceived = Val(vCashReceived) + CN.Execute(sSql).Fields(0).Value
   
     
   sSql = " Select isnull(sum(AdvanceReceived),0) as CreditSale" & vbCrLf _
      + " from BanquetOrder h " & vbCrLf _
      + " where 1=1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.BookingDate = '" & vEntryDate & "'"
   vCashReceived = Val(vCashReceived) + CN.Execute(sSql).Fields(0).Value
   
   sSql = " Select isnull(sum(Received),0) as CreditSale" & vbCrLf _
      + " from BanquetInvoice h " & vbCrLf _
      + " where 1=1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.InvoiceDate = '" & vEntryDate & "'"
   vCashReceived = Val(vCashReceived) + CN.Execute(sSql).Fields(0).Value
   
   
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " FROM CreditVouchers h INNER JOIN CreditVouchersBody b ON h.VoucherNo = B.VoucherNo and h.Storeid = b.Storeid" & vbCrLf _
      + " where BankId is Not null " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.VoucherDate = '" & vEntryDate & "'"
   vBankReceived = CN.Execute(sSql).Fields(0).Value
   
   
   sSql = " Select isnull(sum(Amount),0) as Amount " & vbCrLf _
      + " FROM RecoveryHeader h INNER JOIN RecoveryCustomer b ON h.RecoveryId = B.RecoveryId " & vbCrLf _
      + " where BankMachineiD is Not null  " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.RecoveryDate = '" & vEntryDate & "'"
   vBankReceived = Val(vBankReceived) + CN.Execute(sSql).Fields(0).Value
   
   
   ''''' cash paid on credit Sale Return
   sSql = " Select isnull(sum(cashpaid),0) as SaleReturn" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) amount From SaleReturnBody Group By SID)b on H.SID = B.SID" & vbCrLf _
      + " where Credit = 1 and 1=1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.ReturnDate = '" & vEntryDate & "'"
   
    vCreditSaleReturnPaid = CN.Execute(sSql).Fields(0).Value
    
   ''''' cash paid on Bank Cart Sale Return
   sSql = " Select isnull(sum(cashpaid),0) as SaleReturn" & vbCrLf _
      + " from SaleReturnHeader h inner join (select SID, sum(Amount) amount From SaleReturnBody Group By SID) b on H.SID = B.SID" & vbCrLf _
      + " where BankCard = 1 and 1=1 " & IIf(vUserID = 0, "", " and UserNo = " & vUserID) & " and h.ReturnDate = '" & vEntryDate & "'"

   vCreditSaleReturnPaid = Val(vCreditSaleReturnPaid) + CN.Execute(sSql).Fields(0).Value
   
   vCashAvailable = (Val(vTotalSale) + Val(vRecoveryCustomer) + Val(vCashReceived) + Val(vPettyCash) + Val(vServiceCharges) + Val(vSTax)) - (Val(vBankCardSale) + Val(vCreditSale) + Val(vDiscount) + Val(vSaleReturn) + Val(vPayments))
   vCashAvailable = Val(vCashAvailable) - Val(vCreditSaleReturnPaid)
   
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Public Function SetPost(vUserID As Byte, vEntryDate As Date)
' update sale
   sSql = " Update SaleHeader set isposted = 1 where userno = " & vUserID & " and BillDate = '" & vEntryDate & "'"
   CN.Execute sSql
   
   sSql = " Update ServiceHeader set isposted = 1 where userno = " & vUserID & " and BillDate = '" & vEntryDate & "'"
   CN.Execute sSql
   
   ' update Sale Return
   sSql = " Update SaleReturnHeader set isposted = 1 where userno = " & vUserID & " and ReturnDate = '" & vEntryDate & "'"
   CN.Execute sSql
   
   ' update Replacement
   sSql = " Update ReplacementHeader set isposted = 1 where userno = " & vUserID & " and ReplaceDate = '" & vEntryDate & "'"
   CN.Execute sSql
   
   ' update User Closing
   sSql = " Update UserClosingHeader set isposted = 1 where userno = " & vUserID & " and EntryDate = '" & vEntryDate & "'"
   CN.Execute sSql
   
   ' update Recovery Customer
   'sSQL = " Update ReplacementHeader set isposted = 0 where userno = " & vUserID & " and ReplaceDate = '" & vEntryDate & "'"
   'CN.Execute sSQL
   
End Function
Public Function SetNonPost(vUserID As Byte, vEntryDate As Date)
     
   sSql = " Update SaleHeader set isposted = 0 where userno = " & vUserID & " and BillDate = '" & vEntryDate & "'"
   CN.Execute sSql
   
   sSql = " Update ServiceHeader set isposted = 0 where userno = " & vUserID & " and BillDate = '" & vEntryDate & "'"
   CN.Execute sSql
   
   ' update Sale Return
   sSql = " Update SaleReturnHeader set isposted = 0 where userno = " & vUserID & " and ReturnDate = '" & vEntryDate & "'"
   CN.Execute sSql
   
   ' update Replacement
   sSql = " Update ReplacementHeader set isposted = 0 where userno = " & vUserID & " and ReplaceDate = '" & vEntryDate & "'"
   CN.Execute sSql
   
   ' update User Closing
   sSql = " Update UserClosingHeader set isposted = 0 where userno = " & vUserID & " and EntryDate = '" & vEntryDate & "'"
   CN.Execute sSql
  
 
End Function
