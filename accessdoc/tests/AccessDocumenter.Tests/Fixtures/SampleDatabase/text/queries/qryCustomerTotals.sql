SELECT Customers.CustomerName, Sum(Orders.Total) AS Total
FROM Customers INNER JOIN Orders ON Customers.Id = Orders.CustomerId
GROUP BY Customers.CustomerName;
