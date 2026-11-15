SELECT Orders.Id, Orders.CustomerId, Orders.OrderDate, Orders.Status, Orders.Total
FROM Orders
WHERE Orders.Status <> 'SENT'
ORDER BY Orders.OrderDate;
