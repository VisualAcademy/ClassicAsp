* SQL Server 2000

1. ��ġ/����

2. DB ���� : my_database

3. Table���� :	my_memo

���̺� ����

CREATE TABLE [dbo].[my_memo] 
(
	[Num] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (25) NULL ,
	[Email] [varchar] (100) NULL ,
	[Title] [varchar] (140) NULL ,
	[PostDate] [datetime] NULL ,
)

4. ����� ���� 
	UID : my_database_id
	PWD : my_database_pwd

5. ODBC���� : DSN : my_database_dsn

6. ���ǿ� �⺻ �Խ��� �ٿ� -> my_memo��� ��Ī���� ������͸� ��´�.

7. ASP ���� : http://localhost/memo/memo_list.asp

