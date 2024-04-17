select employeeid,firstname+' '+lastname as Full_name from Employees
where lower(FirstName) like 'robert';

select e.firstname+' '+e.lastname as Full_name,o.*
from employees e join Orders o
on e.EmployeeID=o.EmployeeID
where lower(FirstName) like 'robert' and lower(lastname) like 'king' ;

select * from orders
where EmployeeID =(select EmployeeID from Employees where  concat(firstname,lastname) like 'robertking');

SELECT  * FROM Orders
where EmployeeID=(select EmployeeID from employees where CONCAT(firstname,LastName)='RobertKing');


--8.write a query to get all orders records which belong to employee Robert King