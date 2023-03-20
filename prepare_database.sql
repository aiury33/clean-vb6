
create table Company (IDCompany int IDENTITY(1,1) PRIMARY KEY, Name varchar(255), Wallet money)

insert into Company values ('Market', 0)

create table Product (IDProduct int IDENTITY(1,1) PRIMARY KEY, IDCompany int, Name varchar(255), Cost money, Price money, Quantity float)

insert into Product values (1, 'Potato', 1.25, 2.50, 10)

create table Movements (IDProduct int, OperationIn datetime, Cost money, Price money, Quantity float)


select * from Company
select * from Product
select * from Movements