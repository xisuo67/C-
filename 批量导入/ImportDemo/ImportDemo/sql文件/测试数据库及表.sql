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
     SEX  char(2) check(sex='��' or sex='Ů') not null   
 );  
 insert into students values('1','��ǿ',23,'��');  
 insert into students values('2','����',22,'Ů');  
 insert into students values('5','����',22,'��');
