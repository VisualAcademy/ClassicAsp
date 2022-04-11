--기본형 게시판 작성용 데이터베이스 생성
--Drop Database Basic
Create Database Basic
Go

--Basic데이터베이스로 이동
Use BoardView

--Basic테이블 생성
--Drop Table dbo.Basic
Create Table dbo.Basic
(
     Num Int Identity(1, 1) Not Null Primary Key,  --번호
     Name VarChar(25) Not Null,                    --이름
     Email VarChar(100) Null,                     --이메일     
     Title VarChar(150) Not Null,                  --제목
     PostDate DateTime Default GetDate() Not Null, --작성일     
     PostIP VarChar(15) Not Null,                  --작성IP
     Content Text Not Null,                        --내용
     [Password] VarChar(20) Not Null,                --비밀번호
     ReadCount Int Default 0,                    --조회수
     Encoding VarChar(10) Not Null,              --인코딩(HTML/Text)
     Homepage VarChar(100) Null,                 --홈페이지
     ModifyDate DateTime Null,                   --수정일     
     ModifyIP VarChar(15) Null                   --수정IP
)
Go

--Write 페이지에서 사용되는 구문 예시
Insert Basic(Name,Email,Title,PostDate,PostIP,Content,Password,
ReadCount,Encoding,Homepage) 
Values('홍길동','h@h.com','홍길동다녀갑니다.',
GetDate(),'192.168.0.25', '어쩌구 저쩌구...', '1234', 0,'Text', 
'http://www.h.com/')

--List 페이지에서 사용되는 구문 예시
Select * From Basic Order By Num Desc

--View 페이지에서 사용되는 구문 예시
Select * From Basic Where Num = 1

--Modify 페이지에서 사용되는 구문 예시
Update Basic Set Name = '백두산', Email = 'b@b.com', Title = '백두산이 수정함', Content = '내가 내용을 수정했지롱...', Encoding = 'HTML', Homepage = 'http://www.b.com/', ModifyDate = GetDate(), ModifyIP = '192.168.0.25' Where Num = 1
--Delete
Delete Basic Where Num = 1
--Search
Select * From Basic 
Where 
	Name Like '%이름%' Or Title Like '%제목검색%' Or 
	Content Like '%내용검색%'



