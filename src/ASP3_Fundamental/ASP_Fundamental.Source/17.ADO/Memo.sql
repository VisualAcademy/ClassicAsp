--���� �޸��� ���� ���̺�(Memo)
Create Table dbo.Memo
(
	Num int identity(1, 1) Not Null Primary Key,
	Name VarChar(25) Not Null,
	Email VarChar(50) Not Null,
	Title VarChar(200) Not Null,
	PostDate DateTime Not Null,
	PostIP VarChar(15) Not Null 
)
Go

--Select��
Select * From Memo
--Insert��
Insert Memo Values('ȫ�浿', 'hong@h.com', 'ȫ�浿�Դϴ�.', GetDate(), '127.0.0.1')
--Update��
Update Memo Set Email = 'h@h.com' Where Num = 1
--Delete��
Delete Memo Where Num = 1
