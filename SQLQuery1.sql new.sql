select orderid,productid,unitprice,quantity,quantity*unitprice as Sales_amount
from [Order Details]
where quantity >=20 and quantity*unitprice >=5000
order by Sales_amount desc

select * from Orders
where ShipRegion is not null