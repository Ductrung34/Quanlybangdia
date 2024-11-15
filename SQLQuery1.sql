use QLBangDia
go
CREATE TABLE KHACHHANG
(
	Makhachhang varchar (10) not null,
	Hovaten nvarchar(30)not null,
	Diachi nvarchar(30),
	CMND int not null,
	SĐT int not null,
	Gioitinh nvarchar(10),

)
go
CREATE TABLE PHIEUTHUE
(
	Mabangdia varchar (10) not null ,
	Makhachhang varchar (10) not null,
	Tinhtrang nvarchar(10)not null,
	Ngaythue int,
	Soluong int,

)
go
CREATE TABLE BANGDIA
(
	Tinhtrang nvarchar(10) ,
	Tenbangdia nvarchar(30)not null,
	Mabangdia varchar(10) not null,
	Loaibangdia nvarchar(10)not null,
	Soluong int,
	Dongia int not null,


)
go
CREATE TABLE HOADON
(
	Tinhtrang nvarchar(10)not null,
	Mabangdia varchar(10) not null,
	Makhachhang varchar (10) not null,
	Ngaythue int not null,
	Dongia int,
	
	
	
	


)
go
