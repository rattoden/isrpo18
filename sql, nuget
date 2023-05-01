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

Nuget:
Microsoft.Bcl.AsyncInterfaces
System.Buffers
System.Memory
System.Numerics.Vectors
System.Runtime.CompilerServices.Unsafe
System.Text.Encodings.Web
System.Text.Json
System.Threading.Tasks.Extensions
System.ValueTuple

Сборки:
Microsoft Excel 14.0 Object Library
Microsoft Word 14.0 Object Library


private static IsrpoEntities _context = new IsrpoEntities();

public static IsrpoEntities GetContext()
{
    if (_context == null)
	_context = new IsrpoEntities();
    return _context;
}
