create database Isrpo
go
use Isrpo
create table employees 
(
	id_e nvarchar(20) primary key not null,
	role_e nvarchar(max) not null,
	fio_e nvarchar(max) not null,
	login_e nvarchar(max) not null,
	password_e nvarchar(max) not null,
	last_e nvarchar(max) not null,
	type_e nvarchar(max) not null
)
