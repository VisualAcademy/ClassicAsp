Create Table dbo.Vote
(
	UID Int Identity(1, 1) Not Null Primary Key, --일련번호
	Title VarChar(300) Not Null, --설문제목
	Allow_Multi Bit Not Null, --다중선택여부
	RegDate SmallDateTime Not Null Default(GetDate())--설문등록일
)
Go
Insert Vote Values('관심있는 과목', 0, Default)
Select * From Vote
Create Table dbo.Vote_Item
(
	UID Int Identity(1, 1) Not Null, -- 일련번호
	Vote_ID Int Not Null, -- Vote 테이블의 UID값
	Title VarChar(200) Not Null, -- 각항목의 내용
	Selected Int Not Null Default(0) -- 투표수
)
Go
Insert Vote_Item Values(1, 'C#', 0)
Insert Vote_Item Values(1, 'ASP.NET', 0)
Insert Vote_Item Values(1, 'ASP', 0)
Select * From Vote_Item

-- 새로운 설문등록
Insert Vote Values('밥먹으러 갈까요?', 1, Default)
Insert Vote_Item Values(2, '예', 0)
Insert Vote_Item Values(2, '아니요', 0)
Insert Vote_Item Values(2, '글쎄요', 0)


