select *
from employees
where salary = 17000 and first_name = 'lex';

select *
from employees
where salary < 24000 and salary >9000;
---and
select *
from employees
where salary = 17000 or first_name = 'lex';

select *
from employees
where salary = 9000 or salary = 17000;
---or
select *
from employees
where salary between 9000 and 17000;

select *
from employees where salary between 9000 and 24000 and manager_id =100;
---between and keywords

select first_name +' '+ last_name as Full_name, salary,department_id
from employees
where (salary =24000 or salary =9000) and (department_id= 9 or department_id = 10);
--- or and in same statement

select *
from employees
where first_name like '%r';

select *
from employees
where last_name like '__r%';
---like keyword

select *
from employees
where salary in (4200,9000,17000);
--in keyword

select COUNT(*) as number_of_employees
from employees;
-- to count the employees

select first_name +' '+ last_name as ful_name, manager_id,department_id
from employees
where phone_number is null;
--to know whom within the employees hadn't recorded his phone number
