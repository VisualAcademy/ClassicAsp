--한줄 메모장 관련 테이블(Memo)
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

--Select문
Select * From Memo
--Insert문
Insert Memo Values('홍길동', 'hong@h.com', '홍길동입니다.', GetDate(), '127.0.0.1')
--Update문
Update Memo Set Email = 'h@h.com' Where Num = 1
--Delete문
Delete Memo Where Num = 1
