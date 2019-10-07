create database pms

create table projects(pro_id int primary key identity(1000,1),pro_name varchar(30),pro_head varchar(30),client_name varchar(30),deadline date)
select * from projects 

create table employees(emp_id int primary key identity(100,1),emp_project int default 0,emp_adhaar varchar(16),emp_name varchar(30),emp_mobile varchar(12),emp_mail varchar(30),emp_dep varchar(30))
select * from employees

create table resources(res_project int default null,res_name varchar(30),qty_inuse int default null)
select * from resources 

create table tasks(task_id int primary key identity(100,1),task_project int default null,task_title varchar(50),task_description varchar(200),cur_status varchar(20) )
select * from tasks 
drop table tasks
create table company_res(res_id int primary key identity(1000,1),res_name varchar(30),tot_qty int default null,qty_inuse int default null)
select * from projects
drop table company_res