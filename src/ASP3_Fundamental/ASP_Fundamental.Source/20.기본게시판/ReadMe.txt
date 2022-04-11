* SQL Server 2000

1. 설치/삭제

2. DB 생성 : my_database

3. Table생성 :	my_board

테이블 구조

CREATE TABLE [dbo].[my_board] 
(
	[Num] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (25) NULL ,
	[Email] [varchar] (100) NULL ,
	[Title] [varchar] (140) NULL ,
	[PostDate] [datetime] NULL ,
	[ModifyDate] [datetime] NULL ,
	[ReadCount] [int] NULL ,		/* 글 조회수 */
	[PostIP] [varchar] (15) NULL ,
	[ModifyIP] [varchar] (15) NULL ,
	[Encoding] [varchar] (10) NULL ,   /* HTML로 저장할지, Text로 저장할지 결정 */
	[Passwd] [varchar] (20) NULL ,
	[Content] [text] NULL 
)

4. 사용자 생성 
	UID : my_database_id
	PWD : my_database_pwd

5. ODBC설정 : DSN : my_database_dsn

6. 강의용 기본 게시판 다운 -> my_board라는 별칭으로 가상디렉터리 잡는다.

7. ASP 실행 : http://localhost/my_board/board_list.asp

