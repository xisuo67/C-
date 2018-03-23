if DB_ID('Test') is null
 create database Test;
use Test 
if	exists(select * from dbo.sysobjects where id=object_id('students'))
	drop table students
	go
  create table students  
 (  
    SNO varchar(20) primary key,  
     SNAME varchar(20) ,  
     AGE int,  
     SEX  char(2) check(sex='男' or sex='女') not null   
 );  
 insert into students values('1','李强',23,'男');  
 insert into students values('2','刘丽',22,'女');  
 insert into students values('5','张友',22,'男');
