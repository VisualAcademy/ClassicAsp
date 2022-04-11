* SQL Server 2000

1. 설치/삭제

2. DB 생성 : my_database

3. Table생성 :	my_memo

테이블 구조

CREATE TABLE [dbo].[my_memo] 
(
	[Num] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (25) NULL ,
	[Email] [varchar] (100) NULL ,
	[Title] [varchar] (140) NULL ,
	[PostDate] [datetime] NULL ,
)

4. 사용자 생성 
	UID : my_database_id
	PWD : my_database_pwd

5. ODBC설정 : DSN : my_database_dsn

6. 강의용 기본 게시판 다운 -> my_memo라는 별칭으로 가상디렉터리 잡는다.

7. ASP 실행 : http://localhost/memo/memo_list.asp

