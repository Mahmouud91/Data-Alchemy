

select first_name,salary
from employees
where salary>(select salary
from employees
where employee_id='103')
order by salary;

--1-Write a query to display the first name and salary for all employees who earn more than employee number 103 (Employees table).

select distinct e.first_name,m.manager_id
from employees e join employees m
on m.manager_id=e.employee_id;

select manager_id from employees;

select * from employees
where employee_id in (select manager_id from employees);

select count(*) as [No of managers] from employees
where employee_id in (select manager_id from employees);

--2.write a query to get the total number of employees who are managers.

select count(employee_id) as Non_manager_employee_count from employees
where manager_id<>employee_id;

select count(e.employee_id) as Not_manager_emp_count
from employees e join employees m
on e.manager_id=m.employee_id
where e.employee_id<>m.manager_id;

select * from employees

select * from employees
where employee_id not in (select manager_id from employees where manager_id is not null)


select count(first_name)as No_of_emps from employees
where employee_id not in (select manager_id from employees where manager_id is not null)


--3.write a query to get total number of employees (not managers).

select AVG(salary) as average_salary
from employees;

select * from employees
where salary<(select AVG(salary) as average_salary
from employees);

--4.Write a query to display the employees who earn salary less than the average salary for all employees.

select top(10) * from employees
order by hire_date desc;

--5.Write a query to display the employees who hired recently.

select department_id,department_name,location_id
from departments
where location_id =(select location_id from departments
where department_id=9);

--6.Write a query to display the department number and department name for all departments whose location number is equal to the location number of department number 9 (Departments table).

select e.first_name+' '+e.last_name as Full_name,d.department_name,l.city,c.country_name
from employees e left join departments d
on e.department_id=d.department_id
join locations l
on d.location_id=l.location_id
join countries c
on l.country_id=c.country_id;

select e.first_name+' '+e.last_name as Full_name,d.department_name,l.city,c.country_name
from employees e left join departments d
on e.department_id=d.department_id
join locations l
on d.location_id=l.location_id
join countries c
on l.country_id=c.country_id
where LOWER(e.last_name) like '%a%';

--7.Write a query to display the full name, department name, city, and state province, for all employees whose last name contains the letter a.

