* SQL Server 2000

1. ��ġ/����

2. DB ���� : my_database

3. Table���� :	my_board

���̺� ����

CREATE TABLE [dbo].[my_board] 
(
	[Num] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (25) NULL ,
	[Email] [varchar] (100) NULL ,
	[Title] [varchar] (140) NULL ,
	[PostDate] [datetime] NULL ,
	[ModifyDate] [datetime] NULL ,
	[ReadCount] [int] NULL ,		/* �� ��ȸ�� */
	[PostIP] [varchar] (15) NULL ,
	[ModifyIP] [varchar] (15) NULL ,
	[Encoding] [varchar] (10) NULL ,   /* HTML�� ��������, Text�� �������� ���� */
	[Passwd] [varchar] (20) NULL ,
	[Content] [text] NULL 
)

4. ����� ���� 
	UID : my_database_id
	PWD : my_database_pwd

5. ODBC���� : DSN : my_database_dsn

6. ���ǿ� �⺻ �Խ��� �ٿ� -> my_board��� ��Ī���� ������͸� ��´�.

7. ASP ���� : http://localhost/my_board/board_list.asp

