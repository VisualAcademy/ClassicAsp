Create Table dbo.gBook
(
	UID Int Identity(1, 1) Not Null Primary Key, --일련
	[User_Name] VarChar(20) Not Null, --작성자
	Comment VarChar(4000) Not Null, --내용
	RegDate SmallDateTime Default(GetDate()), --작성일
	User_IP VarChar(15) Not Null, --IP주소
	Pwd VarChar(12) Not Null	--암호
)
Go

Insert gBook Values('홍길동', '안녕', GetDate(), '127.0.0.1', '1234')

Select * From gBook
