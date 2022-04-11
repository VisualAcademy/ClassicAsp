
Create Table dbo.Upload
(
	Num Int Identity(1, 1) Not Null Primary Key,  --번호
	Name VarChar(25) Not Null,                    --이름
	Email VarChar(100) Null,                     --이메일     
	Title VarChar(150) Not Null,                  --제목
	PostDate DateTime Default GetDate() Not Null,   --작성일   
	PostIP VarChar(15) Not Null,                    --작성IP
	Content Text Not Null,                         --내용
	Password VarChar(20) Not Null,                 --비밀번호
	ReadCount Int Default 0,                    --조회수
	Encoding VarChar(10) Not Null,          --인코딩(HTML/Text)
	Homepage VarChar(100) Null,                   --홈페이지
	ModifyDate DateTime Null,                    --수정일     
	ModifyIP VarChar(15) Null,                    --수정IP
	FileName VarChar(255) Null,		--파일명
	FileSize Int Null,			--파일사이즈
	DownCount Int Default 0			--다운횟수
)
Go
