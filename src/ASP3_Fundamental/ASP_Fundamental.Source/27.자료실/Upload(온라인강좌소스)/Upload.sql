use upload
go

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

--예시 데이터
--6.1. 입력 폼(BoardWrite.asp)
--6.2. 입력 처리(BoardWriteProcess.asp)
Insert Upload(Name,Email,Title,PostDate,PostIP,Content, 
	Password,ReadCount,Encoding,Homepage,ModifyDate,
	ModifyIP,FileName,FileSize,DownCount)
Values ('Name','Email','Title',GetDate(),'127.0.0.1','Content', 
	'Password',0,'Text','Homepage',Null,
	Null,'FileName',100,0)
--6.3. 출력(리스트;BoardList.asp)
Select * From Upload Order By Num Desc
--6.4. 세부출력(BoardView.asp)
Update Upload Set ReadCount = ReadCount + 1 Where Num = 1--조회수
Select * From Upload Where Num = 1
--6.5. 수정 폼(BoardModify.asp)
Select * From Upload Where Num = 1
--6.6. 수정 처리(BoardModifyProcess.asp)
Select Password From Upload Where Num = 1 And Password = 'Password'
Update Upload
Set
	Name = 'Name', 
	Email = 'Email',
	Title = 'Title',
	Content = 'Content',
	Encoding = 'HTML',
	Homepage = 'Homepage',
	ModifyDate = GetDate(),
	ModifyIP = '127.0.0.1'
Where Num = 1 And Password = 'Password'
--6.7. 삭제 폼(BoardDelete.asp)
--6.8. 삭제 처리(BoardDeleteProcess.asp)
Select Password From Upload Where Num = 1 And Password = 'Password'
Delete Upload Where Num = 1
--6.9. 검색 폼(BoardSearch.asp)
--6.10. 검색 처리(BoardSearchList.asp)
Select * From Upload 
Where 
	Name Like 'Name' 
	Or 
	Title Like 'Title' 
	Or 
	Content Like 'Content'
Order By Num Desc
--6.11. 다운로드(BoardDown.asp)
Update Upload Set DownCount = DownCount + 1 Where Num = 1

















