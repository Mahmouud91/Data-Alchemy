select c.country_name,l.city,r.region_name,d.department_name,e.first_name+' '+e.last_name as full_name,e.salary
from countries as c
INNER JOIN locations as l
ON c.country_id= l.country_id
INNER JOIN regions as r
ON r.region_id=c.region_id
INNER JOIN departments as d
ON D.location_id=l.location_id
INNER JOIN employees as e
ON e.department_id=d.department_id;
--data from 5 tables

select c.country_name,COUNT(e.first_name) as total_employees,SUM(e.salary) as Total_salary
from countries as c
INNER JOIN locations as l
ON c.country_id= l.country_id
INNER JOIN regions as r
ON r.region_id=c.region_id
INNER JOIN departments as d
ON D.location_id=l.location_id
INNER JOIN employees as e
ON e.department_id=d.department_id
GROUP BY c.country_name;
---each country with total salaries and count of employees