-- 한줄 메모장 테이블 생성
Create Table dbo.Memos
(
	Num Int
		Identity(1, 1) Not Null Primary Key,	-- 번호
	Name VarChar(25) Not Null,					-- 이름
	Email VarChar(50) Null,						-- 이메일
	Title VarChar(150) Not Null,				-- 한줄메모(제목)
	PostDate SmallDateTime Default(GetDate()),	-- 작성일
	PostIP VarChar(15) Null						-- IP주소	
)
Go
