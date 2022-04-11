-- 기본 게시판(Basic)용 테이블 설계
Create Table dbo.Basic
(
	Num Int Identity(1, 1) Not Null Primary Key,--번호
	Name VarChar(25) Not Null,				--이름
	Email VarChar(100) Null, 				--이메일	
	Title VarChar(150) Not Null,				--제목
	PostDate DateTime Default GetDate() Not Null,--작성일	
	PostIP VarChar(15) Not Null,				--작성IP
	Content Text Not Null,					--내용
	Password VarChar(20) Not Null,			--비밀번호
	ReadCount Int Default 0,				--조회수
	Encoding VarChar(10) Not Null,	--인코딩(HTML/Text)
	Homepage VarChar(100) Null,				--홈페이지
	ModifyDate DateTime Null,				--수정일	
	ModifyIP VarChar(15) Null				--수정IP
)
Go
-- 수정 예시문
Update Basic Set Name = '홍', Email = 'h@h.com', Title = '제목', Content = '내용', Encoding = 'Text', Homepage = 'http://', ModifyDate = GetDate(), ModifyIP = '127.0.0.1' Where Num = 2

Select * From Basic Where Num = 1
--입력 예시문
Insert Basic Values('홍길동', 'h@h.com', '안녕', GetDate(), '127.0.0.1', '내용', '1234', 0, 'Text', 'http://', GetDate(), '127.0.0.1')
Go

Update Basic Set Name = 'aaaa', Email = 'aaa', Title = 'aaaaa', Content = 'aaaa', Encoding = 'Text', Homepage = 'aaaa', ModifyDate = GetDate(), ModifyIP = '::1' Where Num = 130
Select * From Basic
Go
